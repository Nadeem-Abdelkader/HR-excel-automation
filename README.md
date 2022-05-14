# HR Excel Automation Script

## Author
- [Nadeem Abdelkader](https://github.com/Nadeem-Abdelkader)

## What Is This?
A Python script that automates HR tasks (mostly generating .xlsx files).

## Set Up and Run

1. Download and install Python 3.10 from <https://www.python.org/downloads/> and make sure to add Python to PATH if you are using Windows
2. Clone or download the git repository
   [here](https://github.com/Nadeem-Abdelkader/HR-excel-automation).
    ```sh
    git clone https://github.com/Nadeem-Abdelkader/HR-excel-automation
    ```
3. Navigate to the cloned local repository
    ```sh
    cd HR-excel-automation
    ```

4. Install required libraries
    ```sh
    pip3 install -r requirements.txt
    ```

5. To start the application, run the following command while inside the cloned local repository:
    ```sh
    python3 main.py [input_file]
    ```
   Example
   ```sh
    python3 main.py "JICDC Attendance 20 April-7 May 2022 V2.xls"
    ```
   Output file will be saved to the current working directory as "Actual Attendance Report.xlsx"

