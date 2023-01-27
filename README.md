# matplotlib-hover-labels

This program is an example of using matplotlib library to make graphs with hover labels and uses PyQt5 backend for the GUI.

It was difficult coming up with hover labels for matplotlib that would work with multiple line graphs that would not cause 
artifacting (where it leaves a label when moving from one line graph to another).  There's not great documentation for doing this sort of thing.   The meat of the hover labels can be found in the show_annation function and the self.cursor1 lines.  The meat of using matplatlib to plot graphs using PyQt5 backend is found in the MplCanvas class, creating a instance of it in a variable and adding as a widget in the initUI for the UI.

The program itself was created for a specific need for reading excel file(s) that contain data to populate the left and right
layouts of the UI.  Then selecting items in the left layout would then populate line graphs on the UI and being able to uncheck/check secondary data in the right layout to change those graphs.  

Admittedly, the program could use some better documentation comments.


![1](https://user-images.githubusercontent.com/123666150/215215127-7daccda4-777b-482a-aab4-6123fcab41cf.PNG)
