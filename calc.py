# TODO:
# - scrolling
# - disable cursor?
# - save should prompt for overwrite
# - cell borders
# - after edit position should stay the same
# - pretty formatting?
# - highlight cells of bad formula to show it's messed up
# - only rjust if cell is a number
# - command line arguments
# - column resize bounds checking

import csv
import curses
import curses.textpad
import re
import sys
from collections import namedtuple

Point = namedtuple('Point', 'x y')

MAX_ROWS = 1000
MAX_COLUMNS = 26

ROW_START = 1

EDIT_X = 50
EDIT_Y = 0

COLUMN_OFFSET = 1
COLUMN_WIDTH = 10

SPREADSHEET = None

single_cell = re.compile(r'([A-Z])(\d+)', flags=re.IGNORECASE)
cell_expansion = re.compile(r'([A-Z]\d+(?::[A-Z]\d+)?,?)', flags=re.IGNORECASE)
formula = re.compile(r'^=(?P<name>[A-Z]+)\((?P<args>.+)\)$', flags=re.IGNORECASE)


class SpreadSheet:
    cells = {}
    filename = None

    hide_overflow = True
    column_width = 8

    position = Point(0, 0)
    sheet_pos = Point(0, 0)

    def __init__(self, stdscr):
        self.stdscr = stdscr
        self.sheet = curses.newpad(MAX_ROWS, (MAX_COLUMNS + 1) * COLUMN_WIDTH)  # A-Z plus column for descriptions

    def handle_input(self, ch):
        if ch in (curses.KEY_DOWN, ord('j')):
            self.chgat(curses.A_NORMAL)
            self.position = self.position._replace(y=min(self.position.y+1, MAX_ROWS - ROW_START - 1))
        elif ch in (curses.KEY_UP, ord('k')):
            self.chgat(curses.A_NORMAL)
            self.position = self.position._replace(y=max(self.position.y-1, 0))
        elif ch in (curses.KEY_LEFT, ord('h')):
            self.chgat(curses.A_NORMAL)
            self.position = self.position._replace(x=max(self.position.x-1, 0))
        elif ch in (curses.KEY_RIGHT, ord('l')):
            self.chgat(curses.A_NORMAL)
            self.position = self.position._replace(x=min(self.position.x+1, MAX_COLUMNS-1))
        elif ch == ord('q'):
            sys.exit(0)
        elif ch == ord('n'):
            self.cells.clear()
        elif ch == ord('-'):
            self.column_width -= 1
        elif ch == ord('+'):
            self.column_width += 1
        elif ch == ord('='):
            self.column_width = 8
        elif ch == ord('H'):
            self.hide_overflow = not self.hide_overflow
        elif ch in (ord('y'), ord('d')):
            cell = self.cells.get(self.position, None)
            if cell is not None:
                self.yank = cell.text
            if ch == ord('d'):
                self.cells.pop(self.position, None)
        elif ch == ord('p'):
            if self.yank is not None:
                self.cells[self.position] = Cell(self.yank)
        elif ch == ord('w'):
            s = self.prompt("SAVE TO:", default=self.filename)
            if s:
                self.save(s)
        elif ch in (curses.KEY_ENTER, 10, ord('i')):
            cell = self.cells.get(self.position, None)

            if cell is None:
                default = ''
            else:
                default = cell.text

            s = self.prompt("EDIT:", default=default)
            if not s:
                self.cells.pop(self.position, None)
            else:
                self.cells[self.position] = Cell(s)

    def save(self, filename):
        self.filename = filename

        expanded_cells = dict(self.cells)

        for position in self.cells:
            # pad out any empty rows
            for y in range(position.y+1):
                new_pos = Point(0, y)
                if new_pos not in expanded_cells:
                    expanded_cells[new_pos] = Cell('')

            # pad out any empty columns
            for x in range(position.x+1):
                new_pos = position._replace(x=x)
                if new_pos not in expanded_cells:
                    expanded_cells[new_pos] = Cell('')

        rows = {}
        for position, cell in expanded_cells.items():
            rows.setdefault(position.y, {})[position.x] = cell.text

        with open(filename, 'w') as f:
            writer = csv.writer(f)

            for y, row in sorted(rows.items()):
                writer.writerow([v for (k, v) in sorted(row.items())])

    def load(self, filename):
        self.filename = filename

        with open(filename, 'rU') as f:
            reader = csv.reader(f)

            for y, fields in enumerate(reader):
                for x, field in enumerate(fields):
                    if not field:
                        continue
                    position = Point(x, y)
                    self.cells[position] = Cell(field)

    def prompt(self, text, default=''):
        def handle_key(key):
            if key == 127:  # convert backspace into CTRL-H
                return ord(curses.ascii.ctrl('h'))
            return key

        def maketextbox(h, w, y, x, value=""):
            nw = curses.newwin(h, w, y, x)
            txtbox = curses.textpad.Textbox(nw, insert_mode=True)

            nw.addstr(0, 0, value)

            self.stdscr.refresh()
            self.stdscr.keypad(1)
            try:
                return txtbox.edit(handle_key).strip()
            except KeyboardInterrupt:
                return None
            finally:
                self.stdscr.keypad(0)

        self.redraw()
        self.stdscr.addstr(0, 0, text)
        self.stdscr.refresh()
        input = maketextbox(1, 40, EDIT_Y, EDIT_X, default)
        return input

    def draw_headings(self):
        char = 'A'

        for i in range(MAX_COLUMNS):
            self.sheet.addstr(0, (i + COLUMN_OFFSET) * COLUMN_WIDTH, char.center(COLUMN_WIDTH))
            char = chr(ord(char) + 1)

        for i, y in enumerate(range(1, MAX_ROWS), 1):
            self.sheet.addstr(y, COLUMN_OFFSET, str(i))

    def draw_data(self):
        for position, cell in self.cells.items():
            adjusted_y = position.y + ROW_START
            adjusted_x = (position.x + COLUMN_OFFSET) * COLUMN_WIDTH
            self.sheet.addstr(adjusted_y, adjusted_x, cell.render().rjust(COLUMN_WIDTH))

    def chgat(self, opts):
        adjusted_y = self.position.y + ROW_START
        adjusted_x = (self.position.x + COLUMN_OFFSET) * COLUMN_WIDTH - COLUMN_OFFSET
        self.sheet.chgat(adjusted_y, adjusted_x, COLUMN_WIDTH+1, opts)

    def redraw(self):
        self.stdscr.clear()
        self.sheet.clear()
        self.draw_headings()
        self.draw_data()
        column = ord('A') + self.position.x
        pos = '({}:{})'.format(chr(column), str(self.position.y+1))
        self.stdscr.addstr(0, 1, pos)
        self.chgat(curses.A_REVERSE)  # highlight current cell
        try:
            self.stdscr.addstr(EDIT_Y, EDIT_X, self.cells[self.position].text)
        except Exception:
            pass
        self.stdscr.refresh()

        max_y, max_x = self.stdscr.getmaxyx()

        self.sheet.refresh(self.sheet_pos.y, self.sheet_pos.x, ROW_START, COLUMN_OFFSET, max_y-ROW_START, max_x-COLUMN_OFFSET)


class Cell(namedtuple('Cell', 'text')):
    __slots__ = ()

    def render(self):
        m = formula.match(self.text)

        if m:
            # looks... formulaic...
            try:
                return self.run_formula(m.group('name'), m.group('args'))
            except Exception:
                pass

        if SPREADSHEET.hide_overflow:
            return self.text[:COLUMN_WIDTH-1]
        return self.text

    def run_formula(self, name, args):
        cells = []

        for cell in cell_expansion.findall(args):
            coords = single_cell.findall(cell)
            col_start, row_start = coords[0]

            if len(coords) == 2:  # it's a range, e.g. A1:A5
                col_end, row_end = coords[1]

                for y in range(int(row_start)-1, int(row_end)):
                    for x in range(ord(col_start.upper()) - ord('A'), ord(col_end.upper()) - ord('A') + 1):
                        p = Point(x, y)
                        if p in SPREADSHEET.cells:
                            cells.append(SPREADSHEET.cells[p])
            else:
                assert len(coords) == 1

                p = Point(
                    ord(col_start.upper()) - ord('A'),
                    int(row_start) - 1,
                )

                if p in SPREADSHEET.cells:
                    cells.append(SPREADSHEET.cells[p])

        return getattr(self, name.lower())(cells)

    def ref(self, cells):
        return cells[0].render()

    def sum(self, cells):
        return str(sum(float(c.render()) for c in cells))

    def sub(self, cells):
        assert len(cells) > 1
        import functools
        import operator
        return '%.2f' % functools.reduce(operator.sub, (float(c.render()) for c in cells))

    def mul(self, cells):
        import functools
        import operator
        return '%.2f' % functools.reduce(operator.mul, (float(c.render()) for c in cells))


def main(stdscr):
    global SPREADSHEET
    assert MAX_COLUMNS <= 26

    stdscr.idlok(True)
    stdscr.scrollok(True)

    SPREADSHEET = SpreadSheet(stdscr)

    if len(sys.argv) > 1:
        SPREADSHEET.load(sys.argv[1])

    while True:
        SPREADSHEET.redraw()

        ch = stdscr.getch()

        SPREADSHEET.handle_input(ch)


if __name__ == '__main__':
    curses.wrapper(main)
