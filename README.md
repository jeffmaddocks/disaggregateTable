# disaggregateTable
Uses [pandas](https://pandas.pydata.org) to iterate through a cross-table in excel, and create an exploded excel file with 1 row for every count in the cross-table. Creating a cross-table is aggregation, this is the opposite: dissaggregation. The downside of course is data fideltiy: cross-tables are built using independent variables, so there is no way to recreate the raw data. In other words, if a cross table is a combination of two independent variables AGE & GENDER, there is no way to take a second cross-table AGE & VACCINATION_STATUS and create a 3rd cross-table GENDER & VACCINATION_STATUS ... because we're working from data that has been aggregated by AGE.

| age | yes | no | 
| --- | --- | --- | 
| 0-10 | 1 | 2 | 
| 11-20 | 3 | 1 | 

... becomes:

| age | isvaccinated | 
| --- | --- | 
| 10-11 | yes | 
| 10-11 | no | 
| 10-11 | no | 
| 11-20 | yes | 
| 11-20 | yes | 
| 11-20 | yes | 
| 11-20 | no | 

See [exampletable.xlsx](exampletable.xlsx) and [exploded.xlsx](exploded.xlsx) for the input/output created by this process. The benefits of automation are:

1. accuracy - we are all fallable, and hand-keying this type of information would risk someone forgetting how many times they've inputted data combinations into a form, or how many rows they've created in excel.

2. time - investing time up front in automation reduces the lift required in subsequent runs of the same process.

In the future, additional automation should be added to handle excel files with many cross-tables. Additional work would be needed to match an existing template.

## Setup

1. Download or clone this project. If you choose to download, setup a working directory first. Alternatively, the "git clone" command will create a directory with the name 'disaggregateTable' that contains this repository:
    ```
    git clone https://github.com/jeffmaddocks/disaggregateTable
    ```

2. Once you have a working directory, setup a virtual environment and activate it  - here's an example in bash:
    ```
    mkdir env
    virtualenv -p python3 ./env
    source env/bin/activate
    ```

3. Install required packages by running pip from the terminal: 
    ```
    pip install -r 'requirements.txt'
    ```

4. Enjoy!
