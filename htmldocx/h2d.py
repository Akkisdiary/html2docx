"""
Make 'span' in tags dict a stack
maybe do the same for all tags in case of unclosed tags?
optionally use bs4 to clean up invalid html?

the idea is that there is a method that converts html files into docx
but also have api methods that let user have more control e.g. so they
can nest calls to something like 'convert_chunk' in loops

user can pass existing document object as arg 
(if they want to manage rest of document themselves)

How to deal with block level style applied over table elements? e.g. text align
"""
import re
import argparse
import io
import os
import urllib.request
from urllib.parse import urlparse
from html.parser import HTMLParser

import docx
import docx.table
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_COLOR, WD_ALIGN_PARAGRAPH

from bs4 import BeautifulSoup

# values in inches
INDENT = 0.25
LIST_INDENT = 0.5
MAX_INDENT = 5.5  # To stop indents going off the page


def get_filename_from_url(url):
    return os.path.basename(urlparse(url).path)


def is_url(url):
    """
    Not to be used for actually validating a url, but in our use case we only 
    care if it's a url or a file path, and they're pretty distinguishable
    """
    parts = urlparse(url)
    return all([parts.scheme, parts.netloc, parts.path])


def fetch_image(url):
    """
    Attempts to fetch an image from a url. 
    If successful returns a bytes object, else returns None

    :return:
    """
    try:
        with urllib.request.urlopen(url) as response:
            # security flaw?
            return io.BytesIO(response.read())
    except urllib.error.URLError:
        return None


def remove_last_occurence(ls, x):
    ls.pop(len(ls) - ls[::-1].index(x) - 1)


def remove_whitespace(string):
    string = re.sub(r'\s*\n\s*', ' ', string)
    return re.sub(r'>\s<', '><', string)


def delete_paragraph(paragraph):
    # https://github.com/python-openxml/python-docx/issues/33#issuecomment-77661907
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


font_styles = {
    'b': {'font-weight': 'bold'},
    'strong': {'font-weight': 'bold'},
    'th': {'font-weight': 'bold'},
    'em': {'font-style': 'italic'},
    'i': {'font-style': 'italic'},
    'u': {'text-decoration': 'underline'},
    's': {'text-decoration': 'line-through'},
}

font_tags = ['b', 'strong', 'em', 'i', 'u', 's', 'sup', 'sub', 'th']
block_tags = ['p', 'div']


class HtmlToDocx(HTMLParser):

    def __init__(self):
        super().__init__()
        self.document = Document()
        self.block = None
        self.current_block = None
        self.run_tags = {}
        self.runs = []
        self.add_column = True

    def set_initial_attrs(self, document=None):
        if document:
            self.doc = document

    def get_cell_html(self, soup):
        # Returns string of td element with opening and closing <td> tags removed
        if soup.find_all():
            return '\n'.join(str(soup).split('\n')[1:-1])
        return str(soup)[4:-5]

    def format_block(self, style):
        if 'text-align' in style:
            align = style['text-align']
            if align == 'center':
                self.block.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align == 'right':
                self.block.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif align == 'justify':
                self.block.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if 'margin-left' in style:
            margin = style['margin-left']
            units = re.sub(r'[0-9]+', '', margin)
            margin = int(re.sub(r'[a-z]+', '', margin))
            if units == 'px':
                self.block.paragraph_format.left_indent = Inches(
                    min(margin // 10 * INDENT, MAX_INDENT))
            # TODO handle non px units

    def add_style_to_run(self, style, run):
        if 'color' in style:
            color = re.sub(r'[a-z()]+', '', style['color'])
            colors = [int(x) for x in color.split(',')]
            run.font.color.rgb = RGBColor(
                *colors)
        if 'background-color' in style:
            color = color = re.sub(r'[a-z()]+', '', style['background-color'])
            colors = [int(x) for x in color.split(',')]
            run.font.highlight_color = WD_COLOR.GRAY_25
        if 'font-size' in style:
            size = re.sub(r'[a-z()]+', '', style['font-size'])
            run.font.size = Pt(int(size))
        if 'font-weight' in style:
            weight = style['font-weight']
            if weight == 'bold':
                run.bold = True
        if 'font-style' in style:
            sty = style['font-style']
            if sty == 'italic':
                run.italic = True
        if 'text-decoration' in style:
            if style['text-decoration'] == 'underline':
                run.underline = True
            if style['text-decoration'] == 'line-through':
                run.font.strike = True

    def parse_dict_string(self, string, separator=';'):
        string = string.replace(' ', '').split(separator)
        parsed_dict = dict([x.split(':') for x in string if ':' in x])
        return parsed_dict

    def handle_li(self):
        # check list stack to determine style and depth
        list_depth = len(self.tags['list'])
        if list_depth:
            list_type = self.tags['list'][-1]
        else:
            list_type = 'ul'  # assign unordered if no tag

        if list_type == 'ol':
            list_style = "List Number"
        else:
            list_style = 'List Bullet'

        self.paragraph = self.doc.add_paragraph(style=list_style)
        self.paragraph.paragraph_format.left_indent = Inches(
            min(list_depth * LIST_INDENT, MAX_INDENT))
        self.paragraph.paragraph_format.line_spacing = 1

    def add_image_to_cell(self, cell, image):
        # python-docx doesn't have method yet for adding images to table cells. For now we use this
        paragraph = cell.add_paragraph()
        run = paragraph.add_run()
        run.add_picture(image)

    def handle_img(self, current_attrs):
        if not self.options['include_images']:
            self.skip = True
            self.skip_tag = 'img'
            return
        src = current_attrs['src']
        # fetch image
        src_is_url = is_url(src)
        if src_is_url:
            try:
                image = fetch_image(src)
            except urllib.error.URLError:
                image = None
        else:
            image = src
        # add image to doc
        if image:
            try:
                if isinstance(self.doc, docx.document.Document):
                    self.doc.add_picture(image)
                else:
                    self.add_image_to_cell(self.doc, image)
            except FileNotFoundError:
                image = None
        if not image:
            if src_is_url:
                self.doc.add_paragraph("<image: %s>" % src)
            else:
                # avoid exposing filepaths in document
                self.doc.add_paragraph("<image: %s>" %
                                       get_filename_from_url(src))
        # add styles?

    def handle_table(self):
        """
        To handle nested tables, we will parse tables manually as follows:
        Get table soup
        Create docx table
        Iterate over soup and fill docx table with new instances of this parser
        Tell HTMLParser to ignore any tags until the corresponding closing table tag
        """
        table_soup = self.tables[self.table_no]
        rows, cols = self.get_table_dimensions(table_soup)
        self.table = self.doc.add_table(rows, cols)
        rows = table_soup.find_all('tr', recursive=False)
        cell_row = 0
        for row in rows:
            cols = row.find_all(['th', 'td'], recursive=False)
            cell_col = 0
            for col in cols:
                cell_html = self.get_cell_html(col)
                if col.name == 'th':
                    cell_html = "<b>%s</b>" % cell_html
                docx_cell = self.table.cell(cell_row, cell_col)
                child_parser = HtmlToDocx()
                child_parser.add_html_to_cell(cell_html, docx_cell)
                cell_col += 1
            cell_row += 1

        # skip all tags until corresponding closing tag
        self.instances_to_skip = len(table_soup.find_all('table'))
        self.skip_tag = 'table'
        self.skip = True
        self.table = None

    def handle_starttag(self, tag, attrs):
        style = dict(attrs).get('style')
        if style:
            style = self.parse_dict_string(style)
        if tag in block_tags:
            self.block = self.document.add_paragraph()
            self.current_block = tag
            if style:
                self.format_block(style)
            return
        if tag in font_tags:
            style = font_styles[tag]
        if tag == 'table':
            self.block = self.document.add_table(0, 0)
            self.current_block = tag
        if tag == 'tr':
            self.current_row = self.block.add_row()
        if tag == 'td' and self.add_column:
            self.block.add_column()
        self.run_tags[tag] = {'style': style, 'runs': []}

    def handle_startendtag(self, tag, attrs):
        # style = dict(attrs).get('style')
        # if style:
        #     style = self.parse_dict_string(style)
        if tag == 'br':
            self.handle_data('\n')

    def handle_data(self, data):
        if self.current_row:
            self.current_row.cells[-1].add_paragraph(data)
            return
        if not self.block:
            self.block = self.doc.add_paragraph()
        run = self.block.add_run(data)
        keys = self.run_tags.keys()
        if len(keys) == 0:
            return
        for tag in keys:
            self.run_tags[tag]['runs'].append(run)

    def handle_endtag(self, tag):
        if tag in block_tags or tag == 'tbody':
            return
        if tag == 'tr':
            self.add_column = False
            return
        for run in self.run_tags[tag]['runs']:
            self.add_style_to_run(self.run_tags[tag]['style'], run)
        self.run_tags.pop(tag)

    def ignore_nested_tables(self, tables_soup):
        """
        Returns array containing only the highest level tables
        Operates on the assumption that bs4 returns child elements immediately after
        the parent element in `find_all`. If this changes in the future, this method will need to be updated

        :return:
        """
        new_tables = []
        nest = 0
        for table in tables_soup:
            if nest:
                nest -= 1
                continue
            new_tables.append(table)
            nest = len(table.find_all('table'))
        return new_tables

    def get_table_dimensions(self, table_soup):
        rows = table_soup.find_all('tr', recursive=False)
        cols = rows[0].find_all(['th', 'td'], recursive=False)
        return len(rows), len(cols)

    def get_tables(self):
        if not hasattr(self, 'soup'):
            self.options['include_tables'] = False
            return
            # find other way to do it, or require this dependency?
        self.tables = self.ignore_nested_tables(self.soup.find_all('table'))
        self.table_no = 0

    def run_process(self, html):
        # if self.options['fix_html'] and BeautifulSoup:
        if BeautifulSoup:
            self.soup = BeautifulSoup(html, 'html.parser')
            html = remove_whitespace(str(self.soup))
        else:
            html = remove_whitespace(html)
        # if self.options['include_tables']:
        #     self.get_tables()
        self.feed(html)

    def add_html_to_cell(self, html, cell):
        if not isinstance(cell, docx.table._Cell):
            raise ValueError('Second argument needs to be a %s' %
                             docx.table._Cell)
        unwanted_paragraph = cell.paragraphs[0]
        delete_paragraph(unwanted_paragraph)
        self.set_initial_attrs(cell)
        self.run_process(html)
        # cells must end with a paragraph or will get message about corrupt file
        # https://stackoverflow.com/a/29287121
        if not self.doc.paragraphs:
            self.doc.add_paragraph('')

    def add_html_to_document(self, html, document):
        if not isinstance(html, str):
            raise ValueError('First argument needs to be a %s' % str)
        elif not isinstance(document, docx.document.Document) and not isinstance(document, docx.table._Cell):
            raise ValueError('Second argument needs to be a %s' %
                             docx.document.Document)
        self.set_initial_attrs(document)
        self.run_process(html)

    def from_file(self, file_path, doc_name=None):
        with open(file_path, 'r') as infile:
            html = infile.read()
        self.run_process(html)
        if not doc_name:
            path, file_name = os.path.split(file_path)
            file_name = file_name.split('.')[0]
            doc_name = '%s/new_%s' % ('.', file_name)
        self.document.save('%s.docx' % doc_name)


if __name__ == '__main__':

    arg_parser = argparse.ArgumentParser(
        description='Convert .html file into .docx file with formatting')
    arg_parser.add_argument(
        'filename_html', help='The .html file to be parsed')
    arg_parser.add_argument(
        'filename_docx',
        nargs='?',
        help='The name of the .docx file to be saved. Default new_docx_file_[filename_html]',
        default=None
    )
    arg_parser.add_argument('--bs', action='store_true',
                            help='Attempt to fix html before parsing. Requires bs4. Default True')

    args = vars(arg_parser.parse_args())
    file_html = args.pop('filename_html')
    html_parser = HtmlToDocx()
    html_parser.parse_html_file(file_html, **args)
