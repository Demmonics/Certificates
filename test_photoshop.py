import win32com.client

ps = win32com.client.Dispatch("Photoshop.Application")
ps.Visible = True
print("Connected")

doc = ps.Open(r"C:\D drive\Certificates\Certificate.psd")
print(f"Opened: {doc.Name}")

doc.Close(2)  # close without saving
print("Success")