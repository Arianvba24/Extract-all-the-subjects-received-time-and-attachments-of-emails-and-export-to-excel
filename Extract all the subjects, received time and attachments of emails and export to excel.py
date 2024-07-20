import win32com.client
import pandas as pd

Outlook = win32com.client.Dispatch("Outlook.Application")
valores = Outlook.GetNamespace("MAPI").GetDefaultFolder(6)

subject_list = []
received_time_list = []
attachments_list = []


for i in valores.Items:
    
    try:

        
        subject = i.Subject
        received_time = i.ReceivedTime
        attachments = i.Attachments
        attachment_names = [attachment.FileName for attachment in attachments]
        subject_list.append(subject)
        received_time_list.append(received_time)
        if len(attachment_names) == 0:
            attachments_list.append("No attachments found")
        else:
            
            attachments_list.append(attachment_names)
             
    except Exception as e:
        print("Error")
        continue

p_values = []
for i in attachments_list:
    if type(i) == str:
        p_values.append(i)
    elif type(i) == list:
        p_value = ",".join(i)
        p_values.append(p_value)
        
p_values
    

time_values = []
for i in received_time_list:
    time_values.append(f"{i.day}/{i.month}/{i.year} {i.hour}:{i.minute}:{i.second}")
    
    
data_values = {
    "Subject": subject_list,
    "Data Reception" : time_values,
    "Attachments" : p_values
    
}

df = pd.DataFrame(data_values)

df["Data Reception"] = pd.to_datetime(df["Data Reception"], format ="%d/%m/%Y %H:%M:%S")
df.to_excel(r"C:\Users\Cash\Documents\pruebas_python\proyectos\outlook\emails.xlsx", index = False)
