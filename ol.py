import os 
import pandas as pd
import smtplib
from email.message import EmailMessage
from getpass import getpass
import win32com.client as win32
import imghdr


# sender_email=input("Enter Your Email ID  ")
# sender_pass=getpass("Enter Your Password ")

df=pd.read_csv("email.csv")
receivers_email=df["email"].values
sub=("Test Mail ")
#attach_files=df["Files to be attached"]
name=df["fname"].values




zipped=zip(receivers_email,name)

for(a,c) in zipped:
    
    # Open the Outlook
    outlook = win32.Dispatch('outlook.application')
    
    # Create the email
    msg = outlook.CreateItem(0)
    files=[(r"pythonpdf.pdf")]
    
    for file in files:
        
        with open(file,'rb') as f:
            
            file_data=f.read()
            file_name=f.name
            
        # msg['From']=sender_email
        msg.To=a
        msg.Subject=sub
        temp_text= open('pycontent.txt','r').read()
        msg.HTMLBody=f"""
        Hi {c},
        <p>Greetings! Hope you are doing great!  </p> 
        <br>          

        <p>Further to our conversation in Linkedin Sales navigator message,</p>
        <br> 

        <p>We have high energetic technical resources working out of Chennai, India with capabilities in the following.</p>
        <br> <br> 


        <table border="1" class="dataframe" id="tableid">
        <thead>
          <tr style="text-align: center;">
            <th></th>
            <th>Robotic Process Automation</th>
            <th>Intelligent Document Processing</th>
            <th>Process mining</th>
            <th>Data Analytics</th>
            <th>Block chain</th>
            <th>API</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Tools</td>
            <td>UiPath Automation Anywhere Blueprism Power Automate Robocorp, Python</td>
            <td>Abby Flexicapture Vantage Uipath Document Understanding Automation Hero iQBot</td>
            <td>Celonis Skan</td>
            <td>PowerBI Alterix Table</td>
            <td>Hyperledger Fabric</td>
            <td>JAVA MERN stack</td>
          </tr>
          <tr>
            <td>Roles</td>
            <td>Business Analyst Solution Architect Developers Service Engineers(support)</td>
            <td>Business Analyst Solution Architect Developers Service Engineers(support)</td>
            <td>Process Analyst Data Engineer</td>
            <td>Data Analyst</td>
            <td>Blockchain developer</td>
            <td>Full stack developer</td>
          </tr>
        </tbody>
        </table>
        <br>

        <p><b>Delivery Excellence:</b></p>
        <br>

        <table border="1" class="dataframe" id="tableid">
        <thead>
          <tr style="text-align: center;">
            <th>BOTs built</th>
            <th>BOTs support</th>
            <th>API developed</th>
            <th>Blockchain project</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>25</td>
            <td>30</td>
            <td>50</td>
            <td>3</td>
          </tr>
        </tbody>
        </table>
        <br>

        <p>We assure the outcome with cost optimisation, on time delivery with quality and inclusive culture. We commit for the following values and goals</p>
        <br>
        <table border="1" class="dataframe" id="tableid">
          <thead>
            <tr style="text-align: center">
              <th>Values(3R)</th>
              <th>Goals(4S)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>1.Responsive-Sense and respond to situation<br>2.Responsible Ownership up to closure<br>3.Resilient Focus on long term sustainable solutions</td>
              <td>1.Stable<br>2.Secure<br>3.Scalable<br>4.Speed</td>
            </tr>
          </tbody>
        </table>
        <br>

        <p>We propose the following engagement model. Looking forward to collaborate with you to transform the workforce from Human workforce to hybrid(Human workers + Digital workers)</p>
        <br>
        <table border="1" class="dataframe" id="tableid">
        <thead>
          <tr style="text-align: center;">
            <th>Time and Material</th>
            <th>Fixed cost</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>1.We would allocate resources.<br>2.We will bill the resources based on per day rate agreed with you</td>
            <td>A.Requirements gathering/Process discovery – Scope the project for time and cost.<br>B.Development/Testing/UAT/Production deployment</td>    </tr>
        </tbody>
        </table>
        <br>

        <P>Refer <a href='https://www.mavdero.in/'>Mavdero.in</a> and attached Mavdero profile.<br>
        Thanks & Regards<br>
        Mathan<br>
        CEO<br>
        Mavdero Techservices Pvt Ltd<br>
        Prince infopark,<br>
        81, B Block,5th Floor, <br>
        2nd main road, Ambattur Industrial estate, <br>
        Ambattur, Chennai <br>
        <a href='https://www.mavdero.in/'>Mavdero.in</a><br>
        Harvard Business Review Advisory Council Member<br>
        Mobile no 91 9840721377</p>
        


        
        """
        # msg.Display()

    
        

        msg.Attachments.Add(os.getcwd() +"\\pythonpdf.pdf")
        
        
        msg.Send()
            
print("All mail sent!")

