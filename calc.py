# TODO:
# - scrolling
# - disable cursor?
# - save/load should remember file name
# - save should prompt for overwrite
# - mention current cell coordinates
# - cell borders
# - formula should support lists as well as ranges
# - make =ref(D12) work instead of =ref(D12:D12)
# - make =sub(D12,D15)
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

HIDE_OVERFLOW = True
YANK = None

FILENAME = None


def debug(s):
    with open('debug', 'a') as f:
        f.write(str(s))
        f.write('\n')


class Cell(namedtuple('Cell', 'text')):
    __slots__ = ()

    def render(self):
        m = re.match(r'^=([a-z]+)\(([a-z]+)(\d+)([:,])([a-z]+)(\d+)\)$', self.text, flags=re.IGNORECASE)

        if m:
            # looks... formulaic...
            try:
                return self.run_formula(m)
            except Exception:
                pass

        if HIDE_OVERFLOW:
            return self.text[:COLUMN_WIDTH-1]
        return self.text

    def run_formula(self, m):
        func, col_start, row_start, separator, col_end, row_end = m.groups()

        func = func.lower()

        _cells = []

        if separator == ',':
            p = Point(
                ord(col_start.upper()) - ord('A'),
                int(row_start) - 1,
            )
            if p in cells:
                _cells.append(cells[p])

            p = Point(
                ord(col_end.upper()) - ord('A'),
                int(row_end) - 1,
            )
            if p in cells:
                _cells.append(cells[p])
        else:
            for y in range(int(row_start)-1, int(row_end)):
                for x in range(ord(col_start.upper()) - ord('A'), ord(col_end.upper()) - ord('A') + 1):
                    p = Point(x, y)
                    if p in cells:
                        _cells.append(cells[p])

        return getattr(self, func)(_cells)

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


cells = {}


def main(stdscr):
    assert MAX_COLUMNS <= 26

    stdscr.idlok(True)
    stdscr.scrollok(True)

    if len(sys.argv) > 1:
        path = sys.argv[1]
        global FILENAME
        FILENAME = path

        with open(path, 'rU') as f:
            reader = csv.reader(f)

            for y, fields in enumerate(reader):
                for x, field in enumerate(fields):
                    if not field:
                        continue
                    position = Point(x, y)
                    cells[position] = Cell(field)

    sheet = curses.newpad(MAX_ROWS, (MAX_COLUMNS + 1) * COLUMN_WIDTH)  # A-Z plus column for descriptions

    def prompt(text, default=''):
        def handle_key(key):
            if key == 127:  # convert backspace into CTRL-H
                return ord(curses.ascii.ctrl('h'))
            return key

        def maketextbox(h, w, y, x, value=""):
            nw = curses.newwin(h, w, y, x)
            txtbox = curses.textpad.Textbox(nw, insert_mode=True)

            nw.addstr(0, 0, value)

            stdscr.refresh()
            stdscr.keypad(1)
            try:
                return txtbox.edit(handle_key).strip()
            except KeyboardInterrupt:
                return None
            finally:
                stdscr.keypad(0)

        redraw()
        stdscr.addstr(0, 0, text)
        stdscr.refresh()
        input = maketextbox(1, 40, EDIT_Y, EDIT_X, default)
        return input

    def draw_headings():
        char = 'A'

        for i in range(MAX_COLUMNS):
            sheet.addstr(0, (i + COLUMN_OFFSET) * COLUMN_WIDTH, char.center(COLUMN_WIDTH))
            char = chr(ord(char) + 1)

        for i, y in enumerate(range(1, MAX_ROWS), 1):
            sheet.addstr(y, COLUMN_OFFSET, str(i))

    def draw_data():
        for position, cell in cells.items():
            adjusted_y = position.y + ROW_START
            adjusted_x = (position.x + COLUMN_OFFSET) * COLUMN_WIDTH
            sheet.addstr(adjusted_y, adjusted_x, cell.render().rjust(COLUMN_WIDTH))

    def chgat(opts):
        adjusted_y = position.y + ROW_START
        adjusted_x = (position.x + COLUMN_OFFSET) * COLUMN_WIDTH - COLUMN_OFFSET
        sheet.chgat(adjusted_y, adjusted_x, COLUMN_WIDTH+1, opts)

    def redraw():
        stdscr.clear()
        sheet.clear()
        draw_headings()
        draw_data()
        chgat(curses.A_REVERSE)  # highlight current cell
        try:
            stdscr.addstr(EDIT_Y, EDIT_X, cells[position].text)
        except Exception:
            pass
        stdscr.refresh()
        sheet.refresh(sheet_pos.y, sheet_pos.x, ROW_START, COLUMN_OFFSET, max_y-ROW_START, max_x-COLUMN_OFFSET)

    position = Point(0, 0)
    sheet_pos = Point(0, 0)

    while True:
        max_y, max_x = stdscr.getmaxyx()

        redraw()

        ch = stdscr.getch()

        if ch in (curses.KEY_DOWN, ord('j')):
            chgat(curses.A_NORMAL)
            position = position._replace(y=min(position.y+1, MAX_ROWS - ROW_START - 1))
        elif ch in (curses.KEY_UP, ord('k')):
            chgat(curses.A_NORMAL)
            position = position._replace(y=max(position.y-1, 0))
        elif ch in (curses.KEY_LEFT, ord('h')):
            chgat(curses.A_NORMAL)
            position = position._replace(x=max(position.x-1, 0))
        elif ch in (curses.KEY_RIGHT, ord('l')):
            chgat(curses.A_NORMAL)
            position = position._replace(x=min(position.x+1, MAX_COLUMNS-1))
        elif ch == ord('q'):
            break
        elif ch == ord('n'):
            cells.clear()
        elif ch == ord('-'):
            global COLUMN_WIDTH
            COLUMN_WIDTH -= 1
        elif ch == ord('+'):
            global COLUMN_WIDTH
            COLUMN_WIDTH += 1
        elif ch == ord('='):
            global COLUMN_WIDTH
            COLUMN_WIDTH = 8
        elif ch == ord('H'):
            global HIDE_OVERFLOW
            HIDE_OVERFLOW = not HIDE_OVERFLOW
        elif ch in (ord('y'), ord('d')):
            global YANK
            cell = cells.get(position, None)
            if cell is not None:
                YANK = cell.text
            if ch == ord('d'):
                cells.pop(position, None)
        elif ch == ord('p'):
            if YANK is not None:
                cells[position] = Cell(YANK)
        elif ch == ord('w'):
            s = prompt("SAVE TO:", default=FILENAME)
            if s:
                global FILENAME
                FILENAME = s
                expanded_cells = dict(cells)

                for position in cells:
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

                with open(s, 'w') as f:
                    writer = csv.writer(f)

                    for y, row in sorted(rows.items()):
                        writer.writerow([v for (k, v) in sorted(row.items())])
        elif ch in (curses.KEY_ENTER, 10, ord('i')):
            cell = cells.get(position, None)

            if cell is None:
                default = ''
            else:
                default = cell.text

            s = prompt("EDIT:", default=default)
            if not s:
                cells.pop(position, None)
            else:
                cells[position] = Cell(s)


if __name__ == '__main__':
    curses.wrapper(main)
