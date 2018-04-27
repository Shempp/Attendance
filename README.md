# Attendance

This is a collection of applications to record attendance of students of an educational institution

## Getting Started

For the program to work correctly, you may need the following components:
*	[Microsoft Visual C++ Redistributable 11.0.61030](https://www.microsoft.com/ru-ru/download/details.aspx?id=30679)
* 	[Qt 5.10.0 or higher](https://www1.qt.io/offline-installers/)
* 	[Virtual COM Port Drivers](http://www.ftdichip.com/Drivers/D2XX.htm)

## How to use?

	1. Open the program **Database** and add to the database all students;
	2. Open the file [conf.ini](conf.ini) and enter the necessary settings for the program **Attendance**;
	3. Open the program **Attendance** and conduct student identification with a reader device (you will need to run the program with each new lesson);
	4. Close the program **Attendance**. You will see the created .xls file with the name of the current date and time;
	5. It is necessary to collect all information at the end of the semester. Create a ***.xls*** group file where all information will be stored;
	6. Run the program **Copy** and select all the necessary files from which you can copy the information created by the program **Attendance**;
	7. Finally it remains to fill in the surnames. Open the program **Database**, synchronize the database file ***Database.xls*** and our created file from item 5.

## Built With

* [Microsoft Visual Studio 2012](https://www.microsoft.com/ru-ru/SoftMicrosoft) - programs Attendance and Copy
* [Qt Creator 4.5.0](https://www.qt.io) - program Database

## License

This project is licensed under the GNU Affero General Public License v3.0 License - see the [LICENSE](LICENSE) file for details

The project uses [ExcelFormat Library](https://www.codeproject.com/Articles/42504/ExcelFormat-Library) which is licensed under The Code Project Open License 1.02 - see the [CPOL-1.02.txt](CPOL-1.02.txt) file for details


