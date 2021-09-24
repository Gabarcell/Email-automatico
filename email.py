import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = "pablo_six@live.com"
email.Subject = "Teste do gabriel"
email.HTMLBody = """
<p>EAI Pablo!!!!!!!!</p>
    <p>é nos .</p>

   <p> Funcionou,</p>
    <p>Att,</p>
    <p>Gabriel Código Python</p>
"""
email.Send()




