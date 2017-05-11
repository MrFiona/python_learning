#######################################################################
#
# An example of creating an Excel chart in a chartsheet with Python
# and XlsxWriter.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org

import xlsxwriter

workbook = xlsxwriter.Workbook('chartsheet.xlsx')

# Add a worksheet to hold the data.
worksheet = workbook.add_worksheet()

# Add a chartsheet. A worksheet that only holds a chart.
chartsheet = workbook.add_chartsheet()

# Add a format for the headings.
bold = workbook.add_format({'bold': 1})

# Add the worksheet data that the charts will refer to.
headings = ['Number', 'Batch 1', 'Batch 2']
data = [
    [2, 3, 4, 5, 6, 7],
    [10, 40, 50, 20, 10, 50],
    [30, 60, 70, 50, 40, 30],
]

worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])
worksheet.write_column('C2', data[2])


# Create a new bar chart.
chart1 = workbook.add_chart({'type': 'bar'})

# Configure the first series.
chart1.add_series({
    'name':       '=Sheet1!$B$1',
    'categories': '=Sheet1!$A$2:$A$7',
    'values':     '=Sheet1!$B$2:$B$7',
})

# Configure a second series. Note use of alternative syntax to define ranges.
chart1.add_series({
    'name':       ['Sheet1', 0, 2],
    'categories': ['Sheet1', 1, 0, 6, 0],
    'values':     ['Sheet1', 1, 2, 6, 2],
})

# Add a chart title and some axis labels.
chart1.set_title ({'name': 'Results of sample analysis'})
chart1.set_x_axis({'name': 'Test number'})
chart1.set_y_axis({'name': 'Sample length (mm)'})

# Set an Excel chart style.
chart1.set_style(11)

# Add the chart to the chartsheet.
chartsheet.set_chart(chart1)

workbook.close()

import cProfile
import pstats
import os
def do_cprofile(filename):
    """
    Decorator for function profiling.
    """
    def wrapper(func):
        def profiled_func(*args, **kwargs):
            # Flag for do profiling or not.
            DO_PROF = os.getenv("PROFILING", True)
            if DO_PROF:
                profile = cProfile.Profile()
                profile.enable()
                result = func(*args, **kwargs)
                profile.disable()
                # Sort stat by internal time.
                sortby = "tottime"
                ps = pstats.Stats(profile).sort_stats(sortby)
                ps.dump_stats(filename)
            else:
                result = func(*args, **kwargs)
            return result
        return profiled_func
    return wrapper

class MicroKineticModel():
    @do_cprofile("./mkm_run.prof")
    def run(self, **kwargs):
        print 'hello world!'

m = MicroKineticModel()
m.run()