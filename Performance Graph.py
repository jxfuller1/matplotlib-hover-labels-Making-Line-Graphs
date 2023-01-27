from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QTabWidget, QWidget, QVBoxLayout, QComboBox, QListWidgetItem, QHBoxLayout, \
    QPushButton, QMenuBar, QListWidget, QGridLayout, QMainWindow, QAction, QCheckBox, QLabel
from PyQt5.QtGui import QColor, QBrush
import sys
import matplotlib
import matplotlib.pyplot as plt
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure
import os
import pandas as pd
from calendar import monthrange
import time
import mplcursors

#month for initializing graph just to have something at start of UI before a month is selected
month = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]
month_pos = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]

current_names_selected = []
current_names_selected_pos = []
#look for certain names within excel that i will be iterating
default_processes = ["INSPECT", "NDI", "DIMENSION"]
current_processes_selected = []
check_state_changed = True
excel_object = ""
excel_object_po = ""



#for graph canvas for widget for pyqt5 for the JOBS tab of the GUI
class MplCanvas(FigureCanvasQTAgg):

    def __init__(self, parent=None, dpi=75):
        global month

        self.fig = Figure(dpi=dpi, tight_layout=True)

        self.axes = self.fig.add_subplot(111, xticks=month)

        self.axes.set_xlabel('Days of Month', fontsize=12)
        self.axes.set_ylabel('Inspections Per Day', fontsize=12)

        super(MplCanvas, self).__init__(self.fig)


# for graph canvas for widget for pyqt5 for the POs tab of the GUI
class MplCanvas_pos(FigureCanvasQTAgg):

    def __init__(self, parent=None, dpi=75):
        global month_pos

        self.fig = Figure(dpi=dpi, tight_layout=True)

        self.axes = self.fig.add_subplot(111, xticks=month_pos)

        self.axes.set_xlabel('Days of Month', fontsize=12)
        self.axes.set_ylabel('Inspections Per Day', fontsize=12)

        super(MplCanvas_pos, self).__init__(self.fig)


#main GUI class
class Actions(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()
        self.month_dates = []
        self.month_dates_pos = []


    def initUI(self):

        self.setGeometry(400,200,1160,615)
        self.setFixedSize(self.size())
        self.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
        self.setWindowTitle("Inspection Performance")

        self.sc = MplCanvas(self, dpi=75)
        self.toolbar = NavigationToolbar2QT(self.sc, self)

        self.sc_pos = MplCanvas_pos(self, dpi=75)
        self.toolbar_pos = NavigationToolbar2QT(self.sc_pos, self)

        self.menubar = QMenuBar()
        self.setMenuBar(self.menubar)

        actionFile = self.menubar.addMenu("File")
        actionFile.addSeparator()
        actionFile.addAction("Quit", self.close_window)
        self.menubar.addMenu("Help")

        self._gray = QBrush(QColor(211,211,211))
        self._green = QBrush(Qt.green)


        self.tab_widget = QWidget(self)
        self.tab_widget.setGeometry(0, 25, self.frameGeometry().width(), self.frameGeometry().height()-50)

        self.layout_tabs = QVBoxLayout(self)

        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()

        # Add tabs
        self.tabs.addTab(self.tab1, "JOB Inspections")
        self.tabs.addTab(self.tab2, "PO Inspections")


        # Create first tab for JOBS tab
        self.vertical_layout = QVBoxLayout(self)
        self.vertical_layout.setSpacing(25)

        self.date_selection = QHBoxLayout()
        self.date_selection.setAlignment(Qt.AlignHCenter | Qt.AlignTop)

        #this set layouts for PO tab
        self.vertical_layout_pos = QVBoxLayout(self)
        self.vertical_layout_pos.setSpacing(25)

        self.date_selection_pos = QHBoxLayout()
        self.date_selection_pos.setAlignment(Qt.AlignHCenter | Qt.AlignTop)


#====================combo box for JOBS and populating it================================================================
        self.combobox1 = QComboBox()

        self.combobox1.setMinimumWidth(150)
        self.combobox1.addItem("--Choose Month/Year--")

        self.combobox1.currentTextChanged.connect(self.read_excel_thread)

        inspect_data_dir = "O:\\folder where excel data to read is stored"
        inspect_months = os.listdir(inspect_data_dir)
        inspect_months = [x for x in inspect_months if "~" not in x]
        inspect_months_jobs = [x for x in inspect_months if "JOB" in x]
        inspect_months_jobs_clean = [s.replace(".csv", "") for s in inspect_months_jobs]

#============================this code just for sorting months in list to proper calender================================
        total_months = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                        "October", "November", "December"]

        index_month = []
        for i in inspect_months_jobs_clean:
            for k in total_months:
                if k in i:
                    index = total_months.index(k)
                    index_month.append(str(index) + " " + i)

        k = 0
        while k < len(index_month):
            temp_split = index_month[k].split(" ")
            index_month[k] = str(temp_split[2]) + " " + str(temp_split[0]) + " " + str(temp_split[1]) + " " + str(temp_split[3])
            k +=1

        index_month.sort()

        k = 0
        while k < len(index_month):
            temp_split = index_month[k].split(" ")
            index_month[k] = str(temp_split[2]) + " " + str(temp_split[0]) + " " + str(temp_split[3])
            k +=1
#============================this code just for sorting months in list to proper calender================================


        self.combobox1.addItems(index_month)

# ====================combo box for JOBS and populating it==============================================================

# ====================combo box for POs and populating it================================================================
        self.comboboxpos = QComboBox()

        self.comboboxpos.setMinimumWidth(150)
        self.comboboxpos.addItem("--Choose Month/Year--")

        self.comboboxpos.currentTextChanged.connect(self.read_excel_thread_pos)

        inspect_data_dir = "O:\\folder for raw data in excel is stored"
        inspect_months = os.listdir(inspect_data_dir)
        inspect_months = [x for x in inspect_months if "~" not in x]
        inspect_months_pos = [x for x in inspect_months if "PO" in x]
        inspect_months_pos_clean = [s.replace(".csv", "") for s in inspect_months_pos]

# ============================this code just for sorting months in list to proper calender================================
        total_months = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                        "October", "November", "December"]

        index_month_pos = []
        for i in inspect_months_pos_clean:
            for k in total_months:
                if k in i:
                    index = total_months.index(k)
                    index_month_pos.append(str(index) + " " + i)

        k = 0
        while k < len(index_month_pos):
            temp_split = index_month_pos[k].split(" ")
            index_month_pos[k] = str(temp_split[2]) + " " + str(temp_split[0]) + " " + str(temp_split[1]) + " " + str(
                temp_split[3])
            k += 1

        index_month_pos.sort()

        k = 0
        while k < len(index_month_pos):
            temp_split = index_month_pos[k].split(" ")
            index_month_pos[k] = str(temp_split[2]) + " " + str(temp_split[0]) + " " + str(temp_split[3])
            k += 1
# ============================this code just for sorting months in list to proper calender================================


        self.comboboxpos.addItems(index_month_pos)
# ====================combo box for POs and populating it==============================================================

        #this set for jobs tab
        self.date_selection.addWidget(self.combobox1)
        self.date_selection.setSpacing(100)
        self.vertical_layout.addLayout(self.date_selection)
        self.graphandnames = QHBoxLayout()


        #this set for pos tab
        self.date_selection_pos.addWidget(self.comboboxpos)
        self.vertical_layout_pos.addLayout(self.date_selection_pos)
        self.graphandnames_pos = QHBoxLayout()


#====================for left hand list in UI================================================================
        self.setStyleSheet("""QListWidget{background: rgb(211,211,211);}""")

        #this set for JOBS
        self.listWidget = QListWidget()
        self.listWidget.setFixedSize(150, 450)
        self.listWidget.sizeHintForColumn(0)
        self.listWidget.itemClicked.connect(self.inspector)

        #this set for POs
        self.listWidgetpos = QListWidget()
        self.listWidgetpos.setFixedSize(150, 450)
        self.listWidgetpos.sizeHintForColumn(0)
        self.listWidgetpos.itemClicked.connect(self.inspector_pos)
# ====================for left hand list in UI================================================================

        #this set for JOBS
        self.graphandnames.addWidget(self.listWidget)
        self.graphandnames.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        self.graph = QVBoxLayout()
        self.graph.addWidget(self.toolbar)
        self.graph.addWidget(self.sc)

        self.graphandnames.addLayout(self.graph)
        self.graphandnames.setSpacing(15)


        self.process_widget = QWidget()
        self.process_widget.setMinimumWidth(150)
        self.processes = QVBoxLayout()
        self.processes.setSpacing(0)
        self.processes.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.process_label = QLabel("<b><u>Process Inspections:<B><U>", self)
        self.process_label.adjustSize()
        self.processes.addWidget(self.process_label)
        self.process_widget.setLayout(self.processes)

        self.graphandnames.addWidget(self.process_widget)


        self.vertical_layout.addLayout(self.graphandnames)

        self.tab1.setLayout(self.vertical_layout)

        #this set for POs
        self.graphandnames_pos.addWidget(self.listWidgetpos)
        self.graphandnames_pos.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        self.graph_pos = QVBoxLayout()
        self.graph_pos.addWidget(self.toolbar_pos)
        self.graph_pos.addWidget(self.sc_pos)

        self.graphandnames_pos.addLayout(self.graph_pos)
        self.graphandnames_pos.setSpacing(0)

        self.vertical_layout_pos.addLayout(self.graphandnames_pos)

        self.tab2.setLayout(self.vertical_layout_pos)


        # Add tabs to widget
        self.layout_tabs.addWidget(self.tabs)
        self.tab_widget.setLayout(self.layout_tabs)


        self.show()

    def tabchanged(self):
        pass

    def close_window(self):
        self.close()

    def inspector(self, inspector1):
        global check_state_changed

        month = self.combobox1.currentText()

        #if month selected in combobox, then change selection on left layout to green
        if "--Choose Month/Year--" not in month:
            color = inspector1.background()
            inspector1.setBackground(self._gray if color == self._green else self._green)
            inspector1.setSelected(False)

        color = inspector1.background()

        #if selected and turned green, reset checkmarks in right side layout
        try:
            if color == self._green:
                if "--Choose Month/Year--" not in month:

                    check_state_changed = False
                    for i in reversed(range(self.processes.count())):  # check mark process inspection checkboxes in UI
                        if i != 0:
                            self.processes.itemAt(i).widget().setChecked(False)
                    check_state_changed = True


                    current_names_selected.append(inspector1.text())
                    self.get_data = Inspector_data(inspector1.text(), month)
                    self.get_data.dataChanged.connect(self.onDataChanged)
                    self.get_data.start()
        except:
            pass

        try:
            if color == self._gray:                          #removes legend/line when unselected in list
                if "--Choose Month/Year--" not in month:
                    current_names_selected.remove(inspector1.text())

                    check_state_changed = False

                    if len(current_names_selected) == 0:  # uncheck everything if no items selected so it can populate properly after one is selected
                        current_processes_selected.clear()
                        for i in reversed(range(self.processes.count())):
                            if i != 0:
                                self.processes.itemAt(i).widget().setChecked(False)
                    check_state_changed = True


                    line = self.sc.axes.get_lines()
                    for i in line:                               #find line name and remove from graph
                        if inspector1.text() in str(i):
                            i.remove()

                    self.sc.axes.legend(loc='upper right', frameon=False, fontsize=9)    #redraw legend after removal
                    self.sc.draw()                  #redraw graph after removal
        except:
            pass


    #for POs tab in GUI  for the left layout
    def inspector_pos(self, inspector1):
        global check_state_changed

        month = self.comboboxpos.currentText()

        if "--Choose Month/Year--" not in month:
            color = inspector1.background()
            inspector1.setBackground(self._gray if color == self._green else self._green)
            inspector1.setSelected(False)

        color = inspector1.background()
        if color == self._green:
            if "--Choose Month/Year--" not in month:
                self.get_data = Inspector_data_pos(inspector1.text(), month)
                self.get_data.dataChanged_pos.connect(self.onDataChanged_pos)
                self.get_data.start()

        if color == self._gray:                          #removes legend/line when unselected in list
            line = self.sc_pos.axes.get_lines()

            for i in line:                               #find line name and remove from graph
                if inspector1.text() in str(i):
                    i.remove()

            self.sc_pos.axes.legend(loc='upper right', frameon=False, fontsize=9)    #redraw legend after removal
            self.sc_pos.draw()                  #redraw graph after removal


    def read_excel_thread(self, month):
        global excel_object, current_names_selected
        if "--Choose Month/Year--" not in month:

            excel_path = "O:\\path to csv for raw data to read to get data that you want to populate in graphs\\" + month + ".csv"
            df = pd.read_csv(excel_path)
            excel_object = df.copy()

            current_names_selected.clear()

            self.calc = External(month)
            self.calc.textChanged.connect(self.onTextChanged)
            self.calc.start()

    def read_excel_thread_pos(self, month):
        global excel_object_po
        if "--Choose Month/Year--" not in month:

            excel_path = "O:\\path to csv for raw data to read to get data that you want to populate in graphs\\" + month + ".csv"
            df = pd.read_csv(excel_path)
            excel_object_po = df.copy()

            self.calc = External(month)
            self.calc.textChanged_pos.connect(self.onTextChanged_pos)
            self.calc.start()

    #activates when change detected in combo box for JOBS tab in GUI for repopulating
    def onTextChanged(self, namelist, datelist, processes1):
        global month, check_state_changed

        self.sc.axes.cla()
        self.listWidget.clear()
        namelist.sort()
        processes1.sort()
        self.listWidget.addItems(namelist)

        to_int = []             #convert datelist into integers
        for i in datelist:
            to_int.append(int(i))

        month.clear()
        self.month_dates.clear()
        month = to_int.copy()        #this one for the global for the graph class
        self.month_dates = to_int.copy()   #this one for the subplotting lines

        for i in reversed(range(self.processes.count())):           #delete checkboxes for recreation for processes in the moonth
            self.processes.itemAt(i).widget().deleteLater()

        self.process_label = QLabel("<b><u>Process Inspections:<B><U>", self)
        self.process_label.adjustSize()
        self.processes.addWidget(self.process_label)

        k = 0
        while k < len(processes1):                        #create new checkboxes for process layout on right side of UI
            self.process_id = QCheckBox(processes1[k])
            self.processes.addWidget(self.process_id)
            self.process_id.stateChanged.connect(self.checkprocess)
            k +=1


        self.graph.removeWidget(self.sc)
        self.graph.removeWidget(self.toolbar)
        self.sc = MplCanvas(self, dpi=75)

        self.toolbar = NavigationToolbar2QT(self.sc, self)
        self.graph.addWidget(self.toolbar)
        self.graph.addWidget(self.sc)


    #activates when change detected in combo box for POs tab in GUI for repopulating
    def onTextChanged_pos(self, namelist, datelist):
        global month, month_pos

        self.sc_pos.axes.cla()
        self.listWidgetpos.clear()
        namelist.sort()
        self.listWidgetpos.addItems(namelist)

        to_int = []             #convert datelist into integers
        for i in datelist:
            to_int.append(int(i))

        month_pos.clear()
        self.month_dates_pos.clear()
        month_pos = to_int.copy()        #this one for the global for the graph class
        self.month_dates_pos = to_int.copy()   #this one for the subplotting lines


        self.graph_pos.removeWidget(self.sc_pos)
        self.graph_pos.removeWidget(self.toolbar_pos)
        self.sc_pos = MplCanvas_pos(self, dpi=75)
        self.toolbar_pos = NavigationToolbar2QT(self.sc_pos, self)
        self.graph_pos.addWidget(self.toolbar_pos)
        self.graph_pos.addWidget(self.sc_pos)

    #for plotting when new item selected on left layout to populate to graph
    def onDataChanged(self, values, name):
        global check_state_changed

        try:
            numberof_zeros = values.count(0)           #remove weekends and days off from calculation
            remove_days_zeros = len(month) - numberof_zeros

            total_values = sum(values)
            if total_values != 0:
                average = sum(values) / remove_days_zeros
            else:
                average = 0

            labelname = name + ", Avg.: " + str(int(average)) + " TTL: " + str(total_values)

            check_state_changed = False

            for i in reversed(range(self.processes.count())):  # check mark process inspection checkboxes in UI
                if self.processes.itemAt(i).widget().text() in current_processes_selected:
                    self.processes.itemAt(i).widget().setChecked(True)

            check_state_changed = True

            line = self.sc.axes.get_lines()

            current_lines = []
            for i in line:
                current_lines.append(str(i))


            if len(current_lines) > 0:
                check_name = [idx for idx in current_lines if name in idx]

                if len(check_name) == 0:
                    self.sc.axes.plot(self.month_dates, values, '-o', label=labelname)
                else:
                    for i in line:  # find line name and remove from graph
                        if name in str(i):
                            i.set_ydata(values)
                            i.set_label(labelname)
            else:
                self.sc.axes.plot(self.month_dates, values, '-o', label=labelname)

        except:
            pass

        try:
            self.cursor1.remove()
        except:
            pass

        self.cursor1 = mplcursors.cursor(self.sc.axes.lines, hover=mplcursors.HoverMode.Transient, multiple = False)
        self.cursor1.connect('add', self.show_annation1)

        self.sc.axes.legend(loc='upper right', frameon=False, fontsize=9)
        self.sc.draw()

    #for hover label in graph
    def show_annation1(self, sel):
        a = str(sel.artist)
        b = a[7:].split(",")
        c = b[0]

        xi, yi = sel.target
        xi = int(round(xi))
        xi = xi - 1
        yi = int(round(yi))

        self.lines = self.sc.axes.get_lines()


        for line in self.lines:
            if line == sel.artist:
                if str(yi) in str(line.get_ydata()[xi]):
                    sel.annotation.set_text(f'{c}\nDay:{month[xi]}\nDone:{yi}')
                else:
                    sel.annotation.set_text()

    #for when new item selected on left layout to populate graph for POs tab
    def onDataChanged_pos(self, values, name):

        numberof_zeros = values.count(0)           #remove weekends and days off from calculation
        remove_days_zeros = len(month_pos) - numberof_zeros

        total_values = sum(values)
        if total_values != 0:
            average = sum(values) / remove_days_zeros
        else:
            average = 0

        labelname = name + ", Avg.: " + str(int(average)) + " TTL: " + str(total_values)

        self.sc_pos.axes.plot(self.month_dates_pos, values, '-o', label=labelname)

        self.sc_pos.axes.legend(loc='upper right', frameon=False, fontsize=9)
        self.sc_pos.draw()

        try:
            self.cursor1_pos.remove()
        except:
            pass

        self.cursor1_pos = mplcursors.cursor(self.sc_pos.axes.lines, hover=mplcursors.HoverMode.Transient, multiple = False)
        self.cursor1_pos.connect('add', self.show_annation1_pos)

    #for hover labels in the PO tab
    def show_annation1_pos(self, sel):
        global month_pos
        a = str(sel.artist)
        b = a[7:].split(",")
        c = b[0]

        xi, yi = sel.target
        xi = int(round(xi))
        xi = xi - 1
        yi = int(round(yi))

        self.lines_pos = self.sc_pos.axes.get_lines()


        for line in self.lines_pos:
            if line == sel.artist:
                if str(yi) in str(line.get_ydata()[xi]):
                    sel.annotation.set_text(f'{c}\nDay:{month_pos[xi]}\nDone:{yi}')
                else:
                    sel.annotation.set_text()

    #this is for populating checkmarks in right side layout based on the item selected in left side layout and reading
    #excel to figure out which checkmarks should exist on the right side layout
    def checkprocess(self):
        global check_state_changed, current_processes_selected, current_names_selected, month

        sender = self.sender()


        if check_state_changed == False:
            pass
        else:
            try:
                current_processes_selected.clear()
                for i in reversed(range(self.processes.count())):
                    if i != 0:
                        box_check = self.processes.itemAt(i).widget().checkState()         #reset current processes selected list
                        if box_check == 2:
                            current_processes_selected.append(self.processes.itemAt(i).widget().text())


                month_check = self.combobox1.currentText()


            # initialize total inspects for each date
                k = 0
                while k < len(current_names_selected):
                    total_on_date = []

                    z = 0
                    while z < len(month):           #reset total on date for next inspector
                        total_on_date.append(0)
                        z += 1

                    if len(current_processes_selected) == 0:
                        line = self.sc.axes.get_lines()
                        for i in line:  # find line name and remove from graph
                            if current_names_selected[k] in str(i):
                                i.set_ydata(total_on_date)
                                labelname = current_names_selected[k] + ", Avg. 0, TTL: 0"
                                i.set_label(labelname)
                        self.sc.axes.legend(loc='upper right', frameon=False, fontsize=9)
                        self.sc.draw()


                    else:
                        for index, row in excel_object.iloc[:, 4].iteritems():  # iterate through process column for inspections
                            if "nan" not in str(row):  # remove items that aren't results
                                if str(row).upper() in current_processes_selected:
                                    name = str(excel_object.iloc[index, 7])
                                    if name in current_names_selected[k]:
                                        date_raw = str(excel_object.iloc[index, 6]).split(" ")
                                        date_raw_2 = date_raw[0].split("/")
                                        date = int(date_raw_2[1])
                                        if date in month:
                                            date_index = month.index(date)
                                            total_on_date[date_index] = total_on_date[date_index] + 1

                        numberof_zeros = total_on_date.count(0)  # remove weekends and days off from calculation
                        remove_days_zeros = len(month) - numberof_zeros
                        total_values = sum(total_on_date)

                        if total_values != 0:
                            average = sum(total_on_date) / remove_days_zeros
                        else:
                            average = 0

                        labelname = current_names_selected[k] + ", Avg.: " + str(int(average)) + " TTL: " + str(total_values)

                        line = self.sc.axes.get_lines()
                        for i in line:  # find line name and remove from graph
                            if current_names_selected[k] in str(i):
                                i.set_ydata(total_on_date)
                                i.set_label(labelname)

                        self.sc.axes.legend(loc='upper right', frameon=False, fontsize=9)
                        self.sc.draw()

                    k +=1
            except:
                pass

            try:
                self.cursor1.remove()
            except:
                pass

            #activate hover labels
            self.cursor1 = mplcursors.cursor(self.sc.axes.lines, hover=mplcursors.HoverMode.Transient, multiple = False)
            self.cursor1.connect('add', self.show_annation1)


class External(QThread):
    textChanged = pyqtSignal(list, list, list)
    textChanged_pos = pyqtSignal(list, list)

    def __init__(self, month):
        super().__init__()
        self.month = month

    def run(self):
        global month, month_pos, excel_object, excel_object_po

        #for setting days in month on graph
        total_days = 0
        total_months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        k = 0
        while k < len(total_months):
            if total_months[k] in self.month:
                a = monthrange(2022, k+1)
                total_days = a[1]
            k += 1

        total_days_list = []
        k = 0
        while k < total_days:
            a = k+1
            total_days_list.append(a)
            k+=1

        if "PO" in self.month:
            month_pos = total_days_list.copy()
        else:
            month = total_days_list.copy()



        total_names = []
        total_processes = []



        if "PO" in self.month:
            for index, row in excel_object_po.iloc[:, 4].iteritems():
                if "nan" not in str(row):
                    name = str(excel_object_po.iloc[index, 4])
                    if name not in total_names:
                        total_names.append(str(name))
            self.textChanged_pos.emit(total_names, total_days_list)

        else:
            #for finding all inspector names to populate for excel to populate on left side of GUI
            for index, row in excel_object.iloc[:, 4].iteritems():
                if "nan" not in str(row):
                    if "INSPECT" in str(row).upper() or "NDI" in str(row).upper() or "DIMENSION" in str(row).upper():
                        name = str(excel_object.iloc[index, 7])
                        process = str(excel_object.iloc[index, 4])
                        if name not in total_names:
                            total_names.append(str(name))
                        if process not in total_processes:
                            total_processes.append(str(process))
            self.textChanged.emit(total_names, total_days_list, total_processes)



class Inspector_data(QThread):
    dataChanged = pyqtSignal(list, str)

    def __init__(self, inspector, month_name):
        super().__init__()
        self.inspector = inspector
        self.month = month_name


    def run(self):
        global month, current_processes_selected, default_processes, current_names_selected

        try:
            current_processes_selected.clear()
            k = 0
            while k < len(current_names_selected):

                total_on_date = []
                z = 0
                while z < len(month):
                    total_on_date.append(0)
                    z += 1

                for index, row in excel_object.iloc[:, 4].iteritems():                # iterate through column in excel looking for what i want
                    if "nan" not in str(row):                   # remove items that aren't results
                        if any(substring in str(row).upper() for substring in default_processes):
                            name = str(excel_object.iloc[index, 7])
                            if name in current_names_selected[k]:
                                date_raw = str(excel_object.iloc[index, 6]).split(" ")
                                date_raw_2 = date_raw[0].split("/")
                                date = int(date_raw_2[1])
                                if date in month:
                                    date_index = month.index(date)
                                    total_on_date[date_index] = total_on_date[date_index] + 1

                                if str(row).upper() not in current_processes_selected:   #add to current processes selected in right of UI
                                    current_processes_selected.append(str(row).upper())


                self.dataChanged.emit(total_on_date, current_names_selected[k])

                k+=1
        except:
            pass

class Inspector_data_pos(QThread):
    dataChanged_pos = pyqtSignal(list, str)

    def __init__(self, inspector, month):
        super().__init__()
        self.inspector = inspector
        self.month = month

    def run(self):
        global month_pos

        total_on_date = []
        z = 0
        while z < len(month_pos):
            total_on_date.append(0)
            z += 1

        excel_path = "O:\\path to excel to read data to gather for graph\\" + self.month + ".csv"
        df = pd.read_csv(excel_path)

        for index, row in df.iloc[:, 4].iteritems():
            if "nan" not in str(row):
                name = str(df.iloc[index, 4])
                if name in self.inspector:
                    date_raw = str(df.iloc[index, 6]).split(" ")
                    date_raw_2 = date_raw[0].split("/")
                    date = int(date_raw_2[1])
                    if date in month:
                        date_index = month.index(date)
                        total_on_date[date_index] = total_on_date[date_index] + 1


        self.dataChanged_pos.emit(total_on_date, self.inspector)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Actions()
    window.show()
    sys.exit(app.exec_())