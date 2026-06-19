This project runs best on a virtual enviroment. To install one in your system, run these cmds in your project directory

Creating the virtual environment : python -m venv .venv
Activating the venv : .venv\Scripts\Activate or .venv\Scripts\activate.bat or .venv\Scripts\Activate.ps1
(there will be a (.venv) on your terminal prompt now)
Downloading the dependecies : pip install -r requirements.txt

Everytime you want to run the site, open terminal and activate the venv and
Run the app.py : python app.py or py app.py
You will see a url. Paste that into a browser to access the site

-------Overview-------
The intended workflow for all the tools are : the html file contains almost no backend. All the calculations and backend proccesses are done by each of the tool's respective python file. Formulaes and lookup tables are present in the respective tool's excel sheets. 

Details of each directory :
/templates contain the flask base (elements common to all pages), and html files for each page
/static contains /css with the styles and /images with the svg and other images used. Also has the word report format
/excel-templates contain the excel sheets (and the corresponding pdf's used for reference) with the formulae, look-up tables and output formats for each tool
/tool_logic contains the py script for each tool
/_pychache_ is a python thing that auto creates when you run app.py
app.py contains all the routing for the files and is what you run to access the site
README.txt is info regarding the repo
requirements.txt contains all the dependencies to be installed

-------Notes-------
1. Runs on a created virtual environment 
3. Exported Excel sheet shows multiple pop ups - click allow for all of them
4. It will NOT run if the respective tool's excel template is kept opened
5. I THINK section_modulus tool has some discrepency and doesn't folloe the intended workflow despite yielding results.

