Hello VB6 DB Developers,
This is a way of expressing my sincerest thanks (to all Developers who contributed and continously
sharing their ideas to VB community) by providing at least one sample appplication that would give an idea
to Beginners as well as in the Intermediate level of Database programming of what an N-Tier application is all
about. This is a transaction oriented-application (see your SQl Profiler)capable of accessing database through stored procedures and without single
Adhoc queries involved. Business rules were implemented through the middle layer which you can compile as in process dll at anytime
and used it by setting the reference and instantiate to exposed its objects. An option to change your connection has been
included during logon to allow user to change database connection. All exeptions are handled by error trapping code so don't
worry about unexpected closing of application. I have also added some crystal report functionality and show how to use
TTX files(Field Definitions Only) on stored procedure .

Brief Description of the Application

This is a fully funtional application developed to monitor production output of a certain machine and operator.

Requirements:

1. SQL Server 2000
2. Crystal Report 9
3. P2smon.dll (CreateFile API Function)
4. iGrid251_75B4A91C.ocx-evaluation (10Tec copyright)
5. LVbuttons.ocx (lavolpe copyright)
6. VB6

Ocx and Dll are provided for your convience.

Application Login
UserID:ALLAN
Password:0123456789

Instructions: 	Set your referenced to P2smon.dll,iGrid251_75B4A91C.ocx and LVbuttons.ocx
		Just put all ocx on c:\Windows\system32 folder
	       	ensure that your computer have SQl Server and Crystal Report installed.
		Attached the two database to your SQL Server using Enterprise Manager.
	       	Create two ODBC connection by selecting SQL Server driver and point your
		Datasource Name to live_environment and test_environment database  and
		name it live and test respectively. 
		
		The Middle Layer (In process-dll) - cDataAccess.dll

		To use in-process dll just unquote the line 'Dim cdbcn As cDataAccess.clsDBAccess located at the second line
		of the declaration section of the forms and unquote 'Set cdbcn = New cDataAccess.clsDBAccess to instatiate the object
		of a class and put a single on 'Dim cdbcn As clsDBAccess located at the first line
		of the declaration section of the forms together with Set cdbcn = New clsDBAccess


I would like to thank Igor Katenev of 10Tec who created iGrid, his wonderfull creation has saved me a lot of time during development
because of its light and easy to use ocx--folks you can try it!, also to Philip Naparan (our Kababayan), to Lavolpe who created lvbuttons
and last but not the least, to PSCode and PSCode fanatics.

Mabuhay ang PSCode (Long Live PSCode)!!!!

Allan V. Pelayo