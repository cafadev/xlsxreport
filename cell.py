class Cell:

    def __init__(self, value, title='', header=False, integer=False, decimal=False,
                 porcentage=False, money=False, money_integer=False, bold=False, empty=False,
                 tr=False, align=None, valign=None, top=None, bottom=None,
                 left=None, right=None, bg=None, fg=None, title_bg=None, title_fg=None,
                 primary_color='#007944'):
        self.primary_color = primary_color
        self.item_table_format = {
            'border': 1,
            'border_color': primary_color
        }

        self.value = value
        self.title = title
        self.style = {} if empty else {**self.item_table_format}
        self.title_bg = title_bg if title_bg is not None else '#fffff'
        self.title_fg = title_fg

        if header:
            self.header_format()

        if integer: self.integer()
        if decimal: self.decimal()
        if porcentage: self.porcentage()
        if money: self.money()
        if money_integer: self.money_integer()
        if bold: self.bold()
        if tr: self.total_row()
        if align is not None: self.align(align)
        if valign is not None: self.valign(valign)
        if bg is not None: self.bg(bg)
        if fg is not None: self.fg(fg)

        self.border(top, right, bottom, left)

    def header_format(self):
        self.style.update({
            'font_color': '#FFFFFF',
            'bg_color': self.primary_color,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'bold': True,
        })

    def bold(self):
        self.style.update({
            'bold': True
        })

    def integer(self):
        self.style.update({
            'num_format': '#,##0',
            'align': 'center',
            'valign': 'vcenter'
        })

    def decimal(self):
        self.style.update({
            'num_format': '#,##0.00',
            'align': 'center',
            'valign': 'vcenter'
        })

    def porcentage(self):
        self.style.update({
            'num_format': '0.00%',
            'align': 'center',
            'valign': 'vcenter'
        })

    def money(self):
        self.style.update({
            'num_format': '_-L* #,##0.00_-;-L* #,##0.00_-;_-L* "-"??_-;_-@_-',
            'valign': 'vcenter'
        })

    def money_integer(self):
        self.style.update({
            'num_format': '_-L* #,##0_-;-L* #,##0_-;_-L* "-"??_-;_-@_-',
            'valign': 'vcenter'
        })

    def total_row(self):
        self.style.update({
            'top': 6
        })

    def align(self, direction):
        self.style.update({
            'align': direction
        })

    def valign(self, direction):
        self.style.update({
            'valign': 'v%s' % direction
        })

    def border(self, top, right, bottom, left):
        if top is not None: self.style.update({'top': top})
        if right is not None: self.style.update({'right': right})
        if bottom is not None: self.style.update({'bottom': bottom})
        if left is not None: self.style.update({'left': left})

    def bg(self, color):
        self.style.update({
            'bg_color': color
        })

    def fg(self, color):
        self.style.update({
            'font_color': color
        })
