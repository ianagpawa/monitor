# Monitor

##### This repo is for my scripts to automate certain tasks involving Excel files, specifically (1) parse data from weekly invoices to create and update a spreadsheet of therapist submissions, and (2) monitor and track if there are missing submissions.


### Quick Start
-Clone the repo: `git clone https://github.com/ianagpawa/monitor.git`


#### Dependencies
These script use `python` and the module `openpyxl` to read and write Excel files.  Install the module with command (`python` and `pip` need to be installed on your system):
```
$   pip install openpyxl
```


#### Running The Scripts
Use the following command to run the main script:
```
python convert.py
```


### File structure
Within the project folder, you will find the following files:

```
monitor/
    ├── Invoices (NOT INCLUDED)/
    |    └── Invoice 170710.xlsx
    ├── .gitignore
    ├── convert.py
    ├── helpers.py
    ├── README.md
    ├── test.py
    ├── test.xlsx
    └── write.xlsx
```

## Creator

**Ian Agpawa**


[Github](https://github.com/ianagpawa)

 agpawaji@gmail.com
