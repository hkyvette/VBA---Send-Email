# VBA---Send-Email
Sending email to brokers --- Daily routine

---Requirements:  
    Send option and future reports of the most recent working day to each broker.
       - In most of the cases, if you send email today, then the date of reports should be yesterday. 
       - If today is Monday, then the most recent working day is last friday.
       - However, when there is public holiday, you need to change the code.

---Logic:
    1. Loop through each folder to find all available reports 
    2. Attach all available attachments
    
