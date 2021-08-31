import json
import requests
from selenium import webdriver
from pywinauto.application import Application
from selenium.webdriver.support.select import Select

chrome = webdriver.Chrome(executable_path=r"C:\Users\SANDEEP NEGI\Documents\ChromeDriver\chromedriver.exe")
chrome.get("https://developer.automationanywhere.com/challenges/automationanywherelabs-employeedatamigration.html?utm_source=challenge_page_success&utm_medium=Developer%20Advocacy&utm_campaign=employee_data_migration")

chrome.maximize_window()



app = Application(backend='uia').start(r"C:\Users\SANDEEP NEGI\Downloads\EmployeeList.exe").connect(title= 'Employee Database',timeout=5)

#app.EmployeeDatabase.print_control_identifiers()

employeeIdEditor = app.EmployeeDatabase.child_window(auto_id="txtEmpId", control_type="Edit").wrapper_object()
searchButton = app.EmployeeDatabase.child_window(title="Search", auto_id="btnSearch", control_type="Button").wrapper_object()
clearButton = app.EmployeeDatabase.child_window(title="Clear", auto_id="btnClear", control_type="Button").wrapper_object()
firstNameTxt = app.EmployeeDatabase.child_window(auto_id="txtFirstName", control_type="Edit").wrapper_object()
lastNameTxt = app.EmployeeDatabase.child_window(auto_id="txtLastName", control_type="Edit").wrapper_object()
emailIdTxt = app.EmployeeDatabase.child_window(auto_id="txtEmailId", control_type="Edit").wrapper_object()
cityTxt = app.EmployeeDatabase.child_window(auto_id="txtCity", control_type="Edit").wrapper_object()
zipTxt = app.EmployeeDatabase.child_window(title="Enter the Emp number", auto_id="txtZip", control_type="Edit").wrapper_object()
stateTxt = app.EmployeeDatabase.child_window(title="Employee Search", auto_id="txtState", control_type="Edit").wrapper_object()
managerTxt = app.EmployeeDatabase.child_window(title="City", auto_id="txtManager", control_type="Edit").wrapper_object()
departmentTxt = app.EmployeeDatabase.child_window(title="Zip Code", auto_id="txtDepartment", control_type="Edit").wrapper_object()
jobTitleTxt = app.EmployeeDatabase.child_window(auto_id="txtJobTitle", control_type="Edit").wrapper_object()

for i in range(10):
    employeeId  = chrome.find_element_by_id("employeeID").get_attribute('value')
    employeeIdEditor.type_keys(employeeId)
    searchButton.click()
    firstName = firstNameTxt.get_value()
    lastName = lastNameTxt.get_value()
    emailId = emailIdTxt.get_value()
    city = cityTxt.get_value()
    zip = zipTxt.get_value()
    state = stateTxt.get_value()
    manager = managerTxt.get_value()
    department = departmentTxt.get_value()
    jobTitle = jobTitleTxt.get_value()
    clearButton.click()

    response = requests.get(f"https://botgames-employee-data-migration-vwsrh7tyda-uc.a.run.app/employees?id={employeeId}")
    jsonOutput = json.loads(response.content.decode('utf-8'))

    chrome.find_element_by_id("firstName").send_keys(firstName)
    chrome.find_element_by_id("lastName").send_keys(lastName)
    chrome.find_element_by_id("phone").send_keys(jsonOutput["phoneNumber"])
    chrome.find_element_by_id("email").send_keys(emailId)
    chrome.find_element_by_id("city").send_keys(city)
    Select(chrome.find_element_by_id("state")).select_by_visible_text(state)
    chrome.find_element_by_id("zip").send_keys(zip)
    chrome.find_element_by_id("title").send_keys(jobTitle)
    Select(chrome.find_element_by_id("department")).select_by_visible_text(department)
    chrome.find_element_by_id("startDate").send_keys(jsonOutput["startDate"])
    chrome.find_element_by_id("manager").send_keys(manager)
    chrome.find_element_by_id("submitButton").click()






