# Train Station Scheduler

This app is created as a Windows Form project in C# programming language.

For the educational purposes for MS Access is used as a database engine.

Unfortunatelly, OleDBConnection to Access have various bugs and programm works perfectely with OLEDB.4.0.

```
connection_string = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=zd.mdb;";
using (OleDbConnection conn = new OleDbConnection(connection_string)){
	conn.Open();
	querryInsert = "INSERT INTO LOGIN([LOGIN], [PASSWORD]) VALUES('" + textBox1.Text + "', '" + textBox2.Text + "');";
	OleDbCommand cmd = new OleDbCommand(querryInsert, conn);
	cmd.ExecuteNonQuery();
}
```

Functionality:
1. Authorization and Registration;
2. Creating a Database of trains, cars, stations, and users;
3. Search for trains;
4. Sorting trains by various parameters;
5. Uploading a list of sorted flights to Excel;
6. Print the ticket.

![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/1.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/2.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/3.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/4.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/5.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/6.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/7.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/8.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/12.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/13.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/28.PNG)
![alt text](https://raw.githubusercontent.com/ramapitecusment/train_schedule_db/master/images/38.PNG)