import os
from dotenv import load_dotenv
import requests
import datetime
from dateutil.parser import parse
import pandas as pd
from jinja2 import Template
import schedule
import time
import urllib.request

load_dotenv()

def connect(host='http://google.com'):
    try:
        urllib.request.urlopen(host)
        return True
    except:
        return False

def getTime():
    msgTime = "Buenos días"
    currentTime = time.strftime("%H:%M:%S",time.localtime())
    print("hora actual es: ",currentTime)
    if currentTime < '12':
        msgTime = "Buenos días"
    if '12' <= currentTime < '18':
        msgTime = "Buenas tardes"
    if currentTime > '18':
        msgTime = "Buenas noches"

    return msgTime

def send_email(data):

    recipient_list = [data["correo"]]

    t = Template("""<html xmlns="http://www.w3.org/1999/xhtml"><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><meta name="viewport" content="width=device-width"><title>{{title}}</title></head><body style="-moz-box-sizing:border-box;-ms-text-size-adjust:100%;-webkit-box-sizing:border-box;-webkit-text-size-adjust:100%;Margin:0;box-sizing:border-box;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0;min-width:100%;padding:0;text-align:left;width:100%!important"><style>@media only screen{html{min-height:100%;background:#f3f3f3}}@media only screen and (max-width:596px){.small-float-center{margin:0 auto!important;float:none!important;text-align:center!important}.small-text-center{text-align:center!important}.small-text-left{text-align:left!important}.small-text-right{text-align:right!important}}@media only screen and (max-width:596px){.hide-for-large{display:block!important;width:auto!important;overflow:visible!important;max-height:none!important;font-size:inherit!important;line-height:inherit!important}}@media only screen and (max-width:596px){table.body table.container .hide-for-large,table.body table.container .row.hide-for-large{display:table!important;width:100%!important}}@media only screen and (max-width:596px){table.body table.container .callout-inner.hide-for-large{display:table-cell!important;width:100%!important}}@media only screen and (max-width:596px){table.body table.container .show-for-large{display:none!important;width:0;mso-hide:all;overflow:hidden}}@media only screen and (max-width:596px){table.body img{width:auto;height:auto}table.body center{min-width:0!important}table.body .container{width:95%!important}table.body .column,table.body .columns{height:auto!important;-moz-box-sizing:border-box;-webkit-box-sizing:border-box;box-sizing:border-box;padding-left:16px!important;padding-right:16px!important}table.body .column .column,table.body .column .columns,table.body .columns .column,table.body .columns .columns{padding-left:0!important;padding-right:0!important}table.body .collapse .column,table.body .collapse .columns{padding-left:0!important;padding-right:0!important}td.small-1,th.small-1{display:inline-block!important;width:8.33333%!important}td.small-2,th.small-2{display:inline-block!important;width:16.66667%!important}td.small-3,th.small-3{display:inline-block!important;width:25%!important}td.small-4,th.small-4{display:inline-block!important;width:33.33333%!important}td.small-5,th.small-5{display:inline-block!important;width:41.66667%!important}td.small-6,th.small-6{display:inline-block!important;width:50%!important}td.small-7,th.small-7{display:inline-block!important;width:58.33333%!important}td.small-8,th.small-8{display:inline-block!important;width:66.66667%!important}td.small-9,th.small-9{display:inline-block!important;width:75%!important}td.small-10,th.small-10{display:inline-block!important;width:83.33333%!important}td.small-11,th.small-11{display:inline-block!important;width:91.66667%!important}td.small-12,th.small-12{display:inline-block!important;width:100%!important}.column td.small-12,.column th.small-12,.columns td.small-12,.columns th.small-12{display:block!important;width:100%!important}table.body td.small-offset-1,table.body th.small-offset-1{margin-left:8.33333%!important;Margin-left:8.33333%!important}table.body td.small-offset-2,table.body th.small-offset-2{margin-left:16.66667%!important;Margin-left:16.66667%!important}table.body td.small-offset-3,table.body th.small-offset-3{margin-left:25%!important;Margin-left:25%!important}table.body td.small-offset-4,table.body th.small-offset-4{margin-left:33.33333%!important;Margin-left:33.33333%!important}table.body td.small-offset-5,table.body th.small-offset-5{margin-left:41.66667%!important;Margin-left:41.66667%!important}table.body td.small-offset-6,table.body th.small-offset-6{margin-left:50%!important;Margin-left:50%!important}table.body td.small-offset-7,table.body th.small-offset-7{margin-left:58.33333%!important;Margin-left:58.33333%!important}table.body td.small-offset-8,table.body th.small-offset-8{margin-left:66.66667%!important;Margin-left:66.66667%!important}table.body td.small-offset-9,table.body th.small-offset-9{margin-left:75%!important;Margin-left:75%!important}table.body td.small-offset-10,table.body th.small-offset-10{margin-left:83.33333%!important;Margin-left:83.33333%!important}table.body td.small-offset-11,table.body th.small-offset-11{margin-left:91.66667%!important;Margin-left:91.66667%!important}table.body table.columns td.expander,table.body table.columns th.expander{display:none!important}table.body .right-text-pad,table.body .text-pad-right{padding-left:10px!important}table.body .left-text-pad,table.body .text-pad-left{padding-right:10px!important}table.menu{width:100%!important}table.menu td,table.menu th{width:auto!important;display:inline-block!important}table.menu.small-vertical td,table.menu.small-vertical th,table.menu.vertical td,table.menu.vertical th{display:block!important}table.menu[align=center]{width:auto!important}table.button.small-expand,table.button.small-expanded{width:100%!important}table.button.small-expand table,table.button.small-expanded table{width:100%}table.button.small-expand table a,table.button.small-expanded table a{text-align:center!important;width:100%!important;padding-left:0!important;padding-right:0!important}table.button.small-expand center,table.button.small-expanded center{min-width:0}}</style><table class="body" data-made-with-foundation="" style="Margin:0;background:#f3f3f3;border-collapse:collapse;border-spacing:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;height:100%;line-height:1.3;margin:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td class="float-center" align="center" valign="top" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0 auto;border-collapse:collapse!important;color:#fff!important;float:none;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0 auto;padding:0;text-align:center;vertical-align:top;word-wrap:break-word"><center data-parsed="" style="min-width:580px;width:100%"><table align="center" class="container float-center" style="Margin:0 auto;background:#2d2d2d;border-collapse:collapse;border-spacing:0;float:none;margin:0 auto;padding:0;text-align:center;vertical-align:top;width:580px"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><table class="spacer" style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td height="16px" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:16px;margin:0;mso-line-height-rule:exactly;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><table class="row" style="border-collapse:collapse;border-spacing:0;display:table;padding:0;position:relative;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td class="small-4 large-4 columns first last" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0 auto;background-color:#0288d1;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0 auto;mso-line-height-rule:exactly;padding:0;padding-bottom:5px;padding-left:16px;padding-right:16px;text-align:left;vertical-align:top;width:177.33px;word-wrap:break-word"></td><td class="small-4 large-4 columns first last" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0 auto;background-color:#03a9f4;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0 auto;mso-line-height-rule:exactly;padding:0;padding-bottom:5px;padding-left:16px;padding-right:16px;text-align:left;vertical-align:top;width:177.33px;word-wrap:break-word"></td><td class="small-4 large-4 columns first last" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0 auto;background-color:#4fc3f7;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0 auto;mso-line-height-rule:exactly;padding:0;padding-bottom:5px;padding-left:16px;padding-right:16px;text-align:left;vertical-align:top;width:177.33px;word-wrap:break-word"></td></tr></tbody></table><br></td></tr></tbody></table><table class="row" style="border-collapse:collapse;border-spacing:0;display:table;padding:0;position:relative;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th class="small-12 large-12 columns first last p-30" style="Margin:0 auto;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0 auto;padding:0;padding-bottom:16px;padding-left:30px;padding-right:30px;text-align:left;width:564px"><table style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left"><img src="http://placehold.it/200x50" alt="" style="-ms-interpolation-mode:bicubic;clear:both;display:block;max-width:100%;outline:0;text-decoration:none;width:auto"></th><th class="expander" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0;padding:0!important;text-align:left;visibility:hidden;width:0"></th></tr></tbody></table></th></tr></tbody></table><img src="http://placehold.it/600x300" alt="" style="-ms-interpolation-mode:bicubic;clear:both;display:block;max-width:100%;outline:0;text-decoration:none;width:auto"><table class="row" style="border-collapse:collapse;border-spacing:0;display:table;padding:0;position:relative;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th class="small-12 large-12 columns first last p-30" style="Margin:0 auto;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0 auto;padding:0;padding-bottom:16px;padding-left:30px;padding-right:30px;text-align:left;width:564px"><table style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left"><br><h2 style="Margin:0;Margin-bottom:10px;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:30px;font-weight:400;line-height:1.3;margin:0;margin-bottom:10px;padding:0;text-align:left;word-wrap:normal"><small style="color:#cacaca;font-size:80%">{{title}}</small></h2><p style="Margin:0;Margin-bottom:10px;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:10pt;font-weight:400;line-height:1.3;margin:0;margin-bottom:10px;padding:0;text-align:left">{{msgTime}} Lorem ipsum dolor sit amet, consectetur adipisicing elit. Nisi repellat, harum. Quas nobis id aut, aspernatur, sequi tempora laborum corporis cum debitis, ullam, dolorem dolore quisquam aperiam! Accusantium, ullam, nesciunt. Lorem ipsum dolor sit amet, consectetur adipisicing elit. Ducimus consequuntur commodi, aut sed, quas quam optio accusantium recusandae nesciunt, architecto veritatis. Voluptatibus sunt esse dolor ipsum voluptates, assumenda quisquam.</p><table class="button large secondary" style="Margin:0 0 16px 0;background:#2d2d2d;border-collapse:collapse;border-spacing:0;margin:0 0 16px 0;padding:0;text-align:left;vertical-align:top;width:auto"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><table style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;background:#777;border:0 solid #777;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><a href="#" style="Margin:0;border:0 solid #777;border-radius:3px;color:#fff!important;display:inline-block;font-family:Helvetica,Arial,sans-serif;font-size:20px;font-weight:700;line-height:1.3;margin:0;padding:10px 20px 10px 20px;text-align:left;text-decoration:none">Click Me!</a></td></tr></tbody></table></td></tr></tbody></table></th><th class="expander" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0;padding:0!important;text-align:left;visibility:hidden;width:0"></th></tr></tbody></table></th></tr></tbody></table><table class="wrapper secondary" align="center" style="background:#2d2d2d;border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td class="wrapper-inner" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><table class="spacer" style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td height="16px" style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:16px;margin:0;mso-line-height-rule:exactly;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><br></td></tr></tbody></table><table class="row" style="border-collapse:collapse;border-spacing:0;display:table;padding:0;position:relative;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th class="small-12 large-12 columns first last p-30" style="Margin:0 auto;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0 auto;padding:0;padding-bottom:16px;padding-left:30px;padding-right:30px;text-align:left;width:564px"><table style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top;width:100%"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left"><div class="button-app"><a href="https://itunes.apple.com/id/" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left;text-decoration:none"><img src="https://s3-ap-southeast-1.amazonaws.com/jojo-receipt/dev-receipts/6bcf1e5a083db2a0f87c42cab007cd27_itunes-app-store-logo.png" height="40" style="-ms-interpolation-mode:bicubic;border:none;clear:both;display:inline;height:40px;max-width:100%;outline:0;text-decoration:none;width:auto"></a><a href="https://play.google.com/store/apps/" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left;text-decoration:none"><img src="https://s3-ap-southeast-1.amazonaws.com/jojo-receipt/dev-receipts/8cf3e1c91db5d1b26b7ed267201e0c5a_google-play-store.png" height="40" style="-ms-interpolation-mode:bicubic;border:none;clear:both;display:inline;height:40px;max-width:100%;outline:0;text-decoration:none;width:auto"></a></div></th></tr></tbody></table></th></tr></tbody></table></td></tr></tbody></table><center data-parsed="" style="min-width:580px;width:100%"><table align="center" class="menu float-center" style="Margin:0 auto;border-collapse:collapse;border-spacing:0;float:none;margin:0 auto;padding:0;text-align:center;vertical-align:top;width:auto!important"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><td style="-moz-hyphens:auto;-webkit-hyphens:auto;Margin:0;border-collapse:collapse!important;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;hyphens:auto;line-height:1.3;margin:0;padding:0;text-align:left;vertical-align:top;word-wrap:break-word"><table style="border-collapse:collapse;border-spacing:0;padding:0;text-align:left;vertical-align:top"><tbody><tr style="padding:0;text-align:left;vertical-align:top"><th class="menu-item float-center" style="Margin:0 auto;color:#fff!important;float:none;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0 auto;padding:10px;padding-right:10px;text-align:center"><a href="#" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left;text-decoration:none">Terms</a></th><th class="menu-item float-center" style="Margin:0 auto;color:#fff!important;float:none;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0 auto;padding:10px;padding-right:10px;text-align:center"><a href="#" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left;text-decoration:none">Privacy</a></th><th class="menu-item float-center" style="Margin:0 auto;color:#fff!important;float:none;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-size:16px;font-weight:400;line-height:1.3;margin:0 auto;padding:10px;padding-right:10px;text-align:center"><a href="#" style="Margin:0;color:#fff!important;font-family:'Gotham Book',Helvetica,Arial,sans-serif;font-weight:400;line-height:1.3;margin:0;padding:0;text-align:left;text-decoration:none">Unsubscribe</a></th></tr></tbody></table></td></tr></tbody></table></center></td></tr></tbody></table></center></td></tr></tbody></table></body></html>""")
    dict = {"msgTime": getTime(),"title": data["nombre"]}
    
    for email in recipient_list:
        print("Enviando el correo a los destinatarios")

        request = requests.post(
            f'https://api.mailgun.net/v3/{os.getenv("DOMAIN_NAME")}/messages',
            auth=("api", f'{os.getenv("API_SECRET_KEY")}'),
            # files = [("test.png", open("files/test.png", "rb", ).read())],
            data={"from": "Lina Magali Dagua Taquinas <info@info.linadagua.online>",
                  "to": email,
                  "subject": "Notificación de asistencia",
                  "text": "Buenas tardes, recuerde que debe asistir a la actividad <ACTIVIDAD> en el lugar <LUGAR>, gracias por su atención",
                  "html": t.render(dict)})

        status = request.status_code
        if (status == 200 or status == 201 or status == 202):
            print("Correo enviado a: ", email)
        else:
            print("Hubo un error al enviar el correo")

def validate(date_text):
        try:
            datetime.date.fromisoformat(date_text)
        except ValueError:
            raise ValueError("Incorrect data format, should be YYYY-MM-DD")

def read_documents():

    archivo_excel = pd.read_excel('docs/cronograma1.xlsx')
    format = "%d-%m-%Y"
    columnas = ['fecha','nombre','correo','motivo']
    df_seleccionados = archivo_excel[columnas]
    print("Obteniendo información de los cronogramas xlsx")
    rows = []
    for data in df_seleccionados:
        row = ''.join(map(str, archivo_excel[data].values))
 
        rows.append(row)  

    rowDate = rows[0]
    if(rowDate != ""):
        fecha = pd.Timestamp(rowDate).date()

        today = datetime.date.today()
        diff = fecha - today
        remainingDays = diff.days
        if(remainingDays == 8 or remainingDays == 7):
            print(f"faltan {remainingDays} días, se enviará el correo")
            
            rows[0] = fecha.strftime("%m/%d/%Y")
            
            data = {
                "fecha": rows[0],
                "nombre": rows[1].capitalize(),
                "correo": rows[2],
                "motivo": rows[3]
            }
 
            send_email(data)
        else:
            print("hoy no se enviará un correo")

def main():

    print("Iniciando...")

    scheduledTime = "09:45:00"

    schedule.every().day.at(scheduledTime).do(read_documents)

    print("Hora programada: ",scheduledTime)
    currentTime = time.strftime("%H:%M:%S",time.localtime())
    if(currentTime > scheduledTime):
        print(f"Son las {currentTime} El servicio se ejecutará mañana")
    if(currentTime <= scheduledTime):
        print("El servicio se ejecutará a las:", scheduledTime)

    


while True:
        
        if(connect()):
            print("Hay internet")
            main()
            schedule.run_pending()
        else:
            print("No hay internet")    
        time.sleep(60)	 	



