# ExcelVBA-Integrated-with-Gmail
This a project where Excel is integrated with Gmail to send Bulk mails with attachment.

 There is a Excel Macros file which have 2 Worksheets i.e.
   `1) HR   :  This Sheet have a data columns i.e. SNo	Name	Email	Title	Company	CC	BCC	Status
               SNo is a Serial Number
               Name is Name of a Person/Employee
               Email is Mail id of that person/Employee
               Title is the designation of the person/Employee
               CC    is the  CC mail id
               BCC   is the BCC mail id
               Status is status which tells mail is send or not.

    2) Email : This sheet have a  mail body that need to be send for concern person. This Sheet also have Gmail Username and password where you need to be fill username & password of your Gmail ID.
                Gmail Password is masked and stored in another cell.

                Subject : Mail body subject. It is in Cell C4.
                Attachment : Mail Attachment. It is in Cell C1 and it contains path of file that need to be attached in mail.
                Signature:  Your Name, mail id, Portfolio address, mobile no....


                # Gmail User name      : It is in Cell I1. You need to fill username of your Gmail Id.
                # Gmail User password  : It is in Cell I2. You need to fill userpassword of your Gmail password. It is masked and stored in another cell BB2.  

                ****## NOTE : The Gmail User password is not that you think. Here are the steps to get password of your Gmail Account.****
                          # 1) Go to your Gmail Account and click Manage your Google Account.
                          # 2) Click  Security : Check " How you sign in to google " have  2-Step Verification is completed. If Completed then go to next step.
                          # 3) Go to Google Account search tab   and search for "App Passwords"
                          # 4) Now New popup screen shows  Enter the password of your Gmail.
                          # 5) Now New popup screen shows Enter your **App Name** , Once you enter the App name then click on **Create** button.
                          # 6) Now New Screen popups that shows **password**, just remember it or save it. This password is used for Gmail Automation.
