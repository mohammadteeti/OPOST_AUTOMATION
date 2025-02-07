# OPOST AUTOMATION

Automate opost reporting of pending follow up , web scrapping project 

## OverView

In opost system the pending shipments must be actively monitored and solved within a short timeframe 

A team of employees frequently monitors pending shipments in the system to ensure fast delivery.
At the end of the day, the pending manager must submit a report on each employee.

### The Problem :

Employee QOS reports are manually generated by reviewing all shipments resolved per employee and storing the data in an Excel file. The back-office manager must then take a random sample and verify response times by manually entering each shipment number into the system.


The reporting process is extremely time-consuming, leading to a backlog of work and significantly hindering overall efficiency.

*__and here comes the solution !__*

### Solution :

A browser automation program using Selenium WebDriver is introduced to streamline the process, reducing manual effort and significantly improving efficiency.

The process that previously took an entire workday is now completed in an hour or less, even for a large number of shipments.

We provide a well-coded, efficient, and user-friendly system as a quick solution for reporting pending employee QOS.


The solution is offered in two working formats:

- Command Line Interface (CLI)  *``on master branch ``*
- Graphical User Interface (GUI) *`` on gui_form_branch ``*
  
This provides flexibility to users with different preferences and technical expertise.



### Features :

- **Browsing Files for Different Employees**

  The program allows the user to easily select the Excel file they wish to work with, eliminating the need for manually entering excessive data.

- **Login to OPOST System**

  The program uses OPOST user credentials to log in, streamlining the login process automatically.

- **Supports Multiple Browsers**

  The program works with both Edge and Chrome, offering a selection mechanism for choosing the browser type. This prevents browser overload, allowing users to perform other tasks on one browser while running the process in another.

- **Version Matching**
  
    The Program checks for browser version and assigns the compatible driver version using driver managers for both Chrome and Edge

- **Wide Range of Shipment Samples**

  The Program allows the user to choose between Full File mode and Sample Mode ,samples can be selected from 5 up to 30 sample per file

- **Logging Mechanism for Tracking Flow and Errors**
  
    A logging window continuously streams the flow of the process, displaying printed outputs, as well as stack, traces for error tracking. This allows for real-time monitoring and troubleshooting of the system's activities. 
- **Start/Stop Features**

    The system allows for multiple start and stop actions during runtime, preventing fatal crashes and ensuring smooth operation even if the process needs to be paused or restarted.
  
- **Configurations File**

    The code relies on a .cfg configuration file to initiate the browser in different modes and on various hosts, enabling multiple running scenarios and providing flexibility for different use cases.
  
- **ٍSimple Notification System**

    The code plays a beep to notify the user about errors in the flow, such as issues with reading shipments or browser connection problems. It also displays a red/green label for quick identification of the issue's status. 




# Image :

![GUI FORM](https://github.com/mohammadteeti/OPOST_AUTOMATION/blob/gui_form_COD/Media/Screen%20Shot.JPG)


# Video :

https://github.com/user-attachments/assets/ee474d1a-61b1-4d68-8fd3-9403f8e7b7ff










