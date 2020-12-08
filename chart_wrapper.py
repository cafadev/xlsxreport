class ChartWrapper:

    def __init__(self, chart, title, scale=None):
        self.scale = {'x_scale': 2.7, 'y_scale': 1.55}
        if scale is not None:
            self.scale.update(scale)

        self.chart = chart
        self.title = title
        self.__set_default_style()

    def __set_default_style(self):
        # self.chart.set_style(1)

        self.chart.set_y_axis({
            'line': {'color': "#ffffff"},
            'major_gridlines': {
                'visible': True,
                'line': {
                    'width': 1,
                    'color': '#F5F5F5'
                }
            },
            'num_font': {
                'color': '#616161'
            }
        })

        self.chart.set_x_axis({
            'major_gridlines': {
                'visible': False,
            }
        })

        self.chart.set_legend({'position': 'top', 'font': {'color': '#616161'}})

        self.chart.set_title ({
            'name': self.title,
            'name_font': {
                'color': '#616161',
                'name': 'Calibri (Cuerpo)',
                'size': 13,
                'bold': False
            },
            'num_font': {
                'bold': False
            }
        })

    def set_table(self):
        self.chart.set_table({
            'show_keys': True,
            'font': {
                'color': '#616161',
            },
        })
