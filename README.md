# Voluntracker
## My capstone project in providing a fully Pythonic standalone GUI for managing volunteers in an organization

What volunteer can do:
- Pull down form responses from a Google Forms output file in Google Sheets
- Use pulled data to track individual user metrics, such as who isn't meeting expectations or who is going above and beyond
- Generate graphs using matplotlib and other Python libraries to visualize key metrics, such as the top 5 performers, bottom 5, organizational average, and most popular volunteering venues
- Populate a template Excel workbook for submitting final results to the parent/oversight organization

Voluntracker is designed to be a totally standalone tkinter GUI application. You do *not* need to install Python
after the project files have been compiled into a standalone EXE using pyinstaller-- it'll run on any machine,
with some restrictions.

Pyinstaller can be finicky to use for this project, since it has trouble with easily identifying Pillow
and all its dependencies. There's a field in the ".spec" file that is generated after first running 
PyInstaller called "Hidden Imports"-- you can add the missing libraries here and kick off a new PyInstaller
build.

This was a great learning experience in data analysis, project management, software development, 
cloud engineering, and overall pushing myself to build a fantastic business solution to a real-world
organization.