# DA-Excel-Dashboard

In this repository here are some interactive excel dashboard templates along with their source file.

a) Sales by chain

b) Sales by category

c) Sales by Manager

d) Sales Total and trend by State

e) Map chart

f) Pie chart

https://www.youtube.com/watch?v=K74_FNnlIF8

1) Create line pivot chart for the month in rows, sales in values and chain in columns pivot table.
2)Remove or hide the year and month filter buttons on line chart since we are going to add slicers to our chart.
Hide all field buttons on the chart.
3)Relocate legend to the top.
4)Hold down shift while arranging the chart elements.

1) To ensure you have same pivot table for sheets, copy the sheet pressing ctrl and shifting the sheet as a new copy and rename it to Category Pivot.
2) Rename pivot table name to CategoryPivot
3) Delete the sales chart.
4) remove months and years.
5) Add category to rows. Chain in columns and sales in values.
6) Create bar chart. Bar charts are good when you have long data labels which ensures they are horizontally formatted and easy to read them.
7) Hide the field buttons and delete the legend. Since all our charts further are going to share the legend created in first chart.
8) Sort Grand Total on pivot table from smallest to largest which makes the sales sorted largest to smallest on the chart. (it works the opposite way)

1) Copy the sheet and rename the sheet. And rename the pivot table to ManagerPivot. Remove category.
2) Get the sales grouped by state and manager in rows.
3) Sort table based on the sales per state in ascending order so that the chart is in descending order.
4) Change Chart title to Sales by manager.

1) For pie chart copy the line pivot.
2) Rename sheet and pivot table to Pie Pivot.
2) Dont like pivot charts as they tend to take a lot of space.
3) Delete chart and remove years and months.
4) Bring the chain in rows.
5) We are going to create a dynamic label so we need the percentage of sales as well.
6) Pop the sales amount in values column again.
7) right click show values as -> % of Grand Total.
8) Rename column as sales % 
9) Get rid of field buttons, legend and the title.
10) resize using ctrl pressed which resizes the pie on all sides equally.

1) Create pivot tables for sparklines.
2) Copy line pivot and rename to sparklines Pivot.
3) Chain in filters, year and months in columns, state in rows, sales in values.
4) Some states are not shown for fashion direct store. So in order to show the values.
Right click -> fieldsettings
On layout and print Tab 
Select show items with no data.
Remove subtotals and grandtotals. From design-> subtotals(remove), grandtotals(remove).
5) Copy paste the pivot table and name it the Next Look pivot.
6) Copy paste and name it Fashions directpivot.

1) For the map chart copy pie pivot sheet. Or any sheet. Rename to Map pivot.
2) Get rid of pie chart.
3) Excel needs to know where the states are. So get the country informationa as well.
4) Get rid of subtotals and grandtotals from design tab.
5) Inorder to make the pivot table compatible with map chart.
We have to show Australia for each state ie in each row.
Right click Australia.
Go to field settings-> Layout and print-> 
select item labels in tabular format.
Select repeat item labels.
6) Ensure the states are sorted.
7) We cant insert map chart on pivot chart. So create a dummy by copying original.
8) So copy paste the pivot tables as values.
9) Then insert map chart based on that.
10) Map chart available from Excel 2016.
Insert->Maps->Map chart.
11) Then point the map to original sales data. And then delete the copied pivot table.
12) Change the color to monochromatic scheme so ti does not use any of orange or blue which are in other charts, just to distinguish it from the rest.
13) rename pivot to Mappivot.

1) Add a new sheet. Name it Dashboard and keep it first in series.
2) Exten the row size and fill with grey.
Write Dashboard name and change font to
Segoe UI Light
Font size : 22
Font color: white
And center it.
3) Now cut paste all the charts on the dashboard.Except sparklines which we need to add manually.

1) For sparklines we cheated by copying the already formatted text columns from the original sheet.
2) Get State while checking for no values using IF.
3)Get sales value from map pivot(cheating) with a check for error using IFError. Convert to currency and no decimal places.
4)Use conditional formatting->solid bar to show the higher and lower values. Usually we dont use both together as it makes the numbers difficult to read. But since here we dont have much room so we sticking them together.
Go to CF->manage rules and change the color of bars to something lighter eg.gray.

1)To insert sparkling select the all chains trend column. Go to insert and select line sparkline.(mini charts placed in single cell).
Seelect the sales values for all months for all states.
To make the values dynamic when any inserts/updates made to the sales value, we add individual sparkline for each state seperately as dynamic does not work for group sparklines. But it would be 8*3 = 24 individual sparklines which would take a lot of time. Hence we select line for the all values and extend it over the last month, for more few months.
For all chains make sparkline black from Design Tab->Black color.

1)Next we are going to add slicers.
There are few ways to add slicers or do anything in excel.
2)One way is to have pivot chart or pivot table have selected. Like sales by chain chart.
If you have excel 2013, right click on field and add a slicer or
you can go on the Insert Tab-> add a slicer. Which gives you list of all the fields and you can select from there.
Add slicers for Category, State and Financial Year.
Resize all slicers as per need.
If you hold down alt, it will snap to the grid. So we can use that shortcut.

Our dashboard is coming together.
Get rid of grid lines(excel cell borders) from View->Uncheck gridlines checkbox.
Keeps the color coding same for all the charts.
Fashions Direct is blue and Next look is Orange.

Lets create a manual legend as all charts are going to use a common legend.
Lets use a rectangle shape and draw it besides the dashboard for store color.
Add a text box besides for adding store name.
Right click and format it for no fill and no outline.
And make the font white.
Get rid of the color shape border.
Roughly align them.
Hold down shift to slect both.
Hold down ctrl and shift and copy them across.
Change the color and name of other store.
We need one more for All Chains. Hence copy paste again using hold down shift so that it aligns with the earlier boxes.
Give it a gradient black fill by going to format tab-> Shape Fill.

Now get rid of the legend in the line chart.
Select sales value in line chart and set the axis so that the units are displayed in thousands which will give a bit more space.
Resize graph and get rid of grid lines.
Repeat fomratting for axis of other charts for consistency.

1)Creating Dynamic Labels for the pie chart.
2)Go to pie pivot and we have to just concatenate the names to show on pie chart.
3)Use & to join and CHAR(10) acts like a line break.
Check video at: https://youtu.be/K74_FNnlIF8?t=2146
4)If the PivotTable Field List pane does not appear click the Analyze tab on the Excel Ribbon, and then click the Field List command.
5)Wrap the GEtpivotdata in Text Functions so that it formats as text with dollar sign.
=A4&CHAR(10)&TEXT(GETPIVOTDATA("Sum of Sales",$A$3,"Chain","Fashions Direct"),"$#,###")&CHAR(10)&TEXT(GETPIVOTDATA("Sales %",$A$3,"Chain","Fashions Direct"),"0%")
6) Apply wrapping so that text appears on next line.
7) Copy that down for getting TExt values for Next Look also.
8)Now we are going to insert it manually by using a text box in pie chart on the dashboard.
Insert->Shapes->Textbox
With the pull handles selected of textbox, click on the formula bar.
Enter [=(select the text from pie pivot sheet)] and press enter.
Change the font color to match the pie.
Align right and font smaller from the Home tab.
Format text box for no fill and no outlines.
9)Hold down ctrl and copy paste.
10)Keep a shade darker then the pie segment as it makes it good to read.
This way we get the dynamic labels.

Now we will connect our slicers with all charts
1)Right click on first slicer and goto report connections.This gets a set of all the pivot tables. Hence its good to name the pivot tables else it gets confusing.
2)Connect 1st slicer to all the pivot tables. Sometimes some pivot tables are not shown in the list. There can be 2 reasons for this. 
For your slicer to control multiple pivot tables, all the pivot tables must use the same source data. In this dataset we same datasource hence its a given.
Sometimes pivot table misses as some tables create a seperate pivot cash even though you have same source data. Everytime you create a pivot table there is a copy of the table stored in excel memory called the pivot cache. Usually there should be only one pivot cache for same source data but sometimes accidently seperate pivot cache is created. Need more understanding here. Please watch the video.
Solution : To recreate any table that is present in the list to get the missing table.
3)For state slicer we connect all tables except the map chart as map chart is not a pivot table.
Other thing is link the map chart to a dynamic named range that picked up the pivot table.
However map chart are new in excel 2016 and they dont yet retain the dynamic name ranges. So we need to work within those limitations until they make any updates.
4)For category slicer apply all pivot tables.
5)Try different values of slicers if all charts are getting updated.
6)Click on clear filter when you want to reset selection.

Linking all the pivots to one source data helps to dynamically update it when there is any new data.
Usually use power query to get the tables dynamically updated. Currently we manually append the august data to our original data.
When you append any data, the table range grows to accomodate any new data.
The range icon i.e triangle extenda down to the end of appended data.
And for the pivot table and pivot charts we just need to click on refresh all.
Goto Data Tab and click on refresh.
Refresh all did not make any changes so i clicked on refresh.

Next thing is to adjust the colors acoording to the branding themes of the stores.
For this use Page Layout->colors.
We can setup our own custom theme that picks up all of our company branding.
We can also choose from some of the built in themes.
We have used parcel. This changes the colors, fonts and layouts are little more spacious.

Group the pie chart objects together. So you can align them together.
From Options->Group or align using Options->Align Slicers to left.

Format->align->Align top the top 3 charts
Format->Check height of any chart to align other charts in alignment with same.
Can use Ctrl+Up arrow to move charts.

Edit Names of slicers to keep them short.
Eg. Financial Year = FY
Category = Cat
Right click->Slicer settings->Slicer captions

Get rid of the column and the row labels. And the sheet tabs and the scroll bars.
On view Tab-> get rif of headings.
File->Options->Advanced->Display
Deslect horizonal and vertical scroll bars and Show sheet tabs.
Remove formula bar from View->uncheck formula bar.

Read the Dashboard Protection Tips given by her in her excel workbook.
Can also embed the dashboard on a webpage 

eBooks and PDF section:
Read Chart Recipe eBook : Helps you decide which chart to use for your data.

Online Excel Dashboard course under Pricing on her site.
