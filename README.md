
![csvexcel](Banner.png)



<p align = "center">
A streamlit application to convert (.xlsx to .csv) or (csv to .xlsx)
<p>

---

<h5 align="center">
    <a href="https://csvexcel.herokuapp.com/"> View streamlit application in Heroku
    </a>
</h5>


---

<h3 align="left">
    Convert Excel to CSV
    <br>
</h3>

1. Upload your `excel` file, if you have multiple worksheets, this application will choose the first one by defult, or you can choose the sheet you want to convert to `csv`
2. Select your columns to convert, you can select all columns.
3. Click on the button [**convert to csv file**]  🚀
4. You can see a preview of generated `csv`
5. Click on the buttom [**Download data as CSV**] ⬇️


<h3 align="left">
    Convert CSV to Excel
    <br>
</h3>

1. Upload your `csv` file
2. Select your columns to convert, you can select all columns.
3. Click on the button [**convert to xlsx file**]  🚀
4. You can see a preview of generated `xlsx`
5. Click on the buttom [**Download data as xlsx**] ⬇️


<h3 align="left">
    Packages used
    <br>
</h3>

* pandas
* streamlit
* io
* datetime
* tabulate
* ssl
* time
* openpyxl

<h3 align="left">
    streamlit siting
    <br>
</h3>

1. Download [csvexcel.py](csvexcel.py), [requirements.txt](requirements.txt), [Procfile](Procfile), [setup.sh](setup.sh) into the **same directory**
2. If you want to change the python file's name, make sur to change it also in [Procfile](Procfile).
3. Use Python terminal (always in the same directory as Python file) and write
```python
streamlit run csvexcel.py
```

---
[Bunner maker](https://banner.godori.dev/)
[Readme.so Editor](https://readme.so/fr/editor)