# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用
###################################################################
#把下方单引号内的内容改为你的应用id                                         #
id=r'81ee469b-fdf6-45bd-b7f9-76eea333f8a4'                         
#把下方单引号内的内容改为你的应用机密                                       #
secret=r'eyJ0eXAiOiJKV1QiLCJub25jZSI6IldmV1FVeFBlYng4V3dzaUNVWUFZb3k2cFFTTGNVS2dMVFFtYWViUVk1QUUiLCJhbGciOiJSUzI1NiIsIng1dCI6Imh1Tjk1SXZQZmVocTM0R3pCRFoxR1hHaXJuTSIsImtpZCI6Imh1Tjk1SXZQZmVocTM0R3pCRFoxR1hHaXJuTSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81ZDg4OWM1OS1mNGYzLTQwZjMtOWE0Mi0zYjJmOTc1NDNhMDQvIiwiaWF0IjoxNTk3NjQ4NzcxLCJuYmYiOjE1OTc2NDg3NzEsImV4cCI6MTU5NzY1MjY3MSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFTUUEyLzhRQUFBQWFUMlN0R1RvcWlVcGM5TyswQmFhcld3YUJUNFdrY3pZaG1QQ0RsZ0ZtY3M9IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJYVVFJIiwiYXBwaWQiOiI4MWVlNDY5Yi1mZGY2LTQ1YmQtYjdmOS03NmVlYTMzM2Y4YTQiLCJhcHBpZGFjciI6IjEiLCJmYW1pbHlfbmFtZSI6ImEiLCJnaXZlbl9uYW1lIjoiYnUiLCJpcGFkZHIiOiI5Ni44LjExOC4xMDgiLCJuYW1lIjoiYnUgYSIsIm9pZCI6IjIzNzRkNmI0LThmOTktNDk3OC1hNWFhLWE1NDRkMGVlZWY3OSIsInBsYXRmIjoiMyIsInB1aWQiOiIxMDAzMjAwMEQ5ODk4NUJGIiwic2NwIjoiRGlyZWN0b3J5LlJlYWQuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWQgRmlsZXMuUmVhZC5BbGwgRmlsZXMuUmVhZFdyaXRlIEZpbGVzLlJlYWRXcml0ZS5BbGwgTWFpbC5SZWFkIE1haWwuUmVhZFdyaXRlIE1haWxib3hTZXR0aW5ncy5SZWFkIE1haWxib3hTZXR0aW5ncy5SZWFkV3JpdGUgU2l0ZXMuUmVhZC5BbGwgU2l0ZXMuUmVhZFdyaXRlLkFsbCBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBVc2VyLlJlYWRXcml0ZS5BbGwgcHJvZmlsZSBvcGVuaWQgZW1haWwiLCJzdWIiOiJqZ2IzM2J6WS02Z0RRZU5jRVZ6aldNVGVtTmJXSWVkY0VPUWtQMG13ek80IiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkFTIiwidGlkIjoiNWQ4ODljNTktZjRmMy00MGYzLTlhNDItM2IyZjk3NTQzYTA0IiwidW5pcXVlX25hbWUiOiJuaWtvNjU2NEBlYzJlYy5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJuaWtvNjU2NEBlYzJlYy5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiJqWkowT2hjNmlFMklXWTByY3pOakFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiXSwieG1zX3N0Ijp7InN1YiI6Inc0VENDdVo4WFk2T0VMN0FPdm8xcUh1b1VpcDVHZE0yZFJyTmlEVTdsdzQifSwieG1zX3RjZHQiOjE1OTc2NDcwODZ9.J4xYOnvhuKo6wGrVGa458Y-SN3NVxVG1scmIPTykc5yR8oRHwK4J5w_89no0cCUFwN1zQmOq285UwpKar8NsYca4CUNBaOIvpqTssa4xNphpL_yLittaqraPP-ANlMGvdYUFnpmF7rRUnTkinksbw5SJ9MhdigYkGvWrflUkuQrEaryuH2VnbkMlQwtveZRprflu3DcuBTXSA_AD7q-s4IDM8kp93qxazvfgGoKd3pDuRIYGCxhdpVmXqtO_Z-Bd3FWaY1uD-PcMkL3M3GxGWNtBMFfwrm-Y64epZfLz39jq1Mvswwe231xGTNu3A8Ya8eGUyvvSFJtbjTP9awD4Kw'                                           
###################################################################


path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':id,
          'client_secret':secret,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 : '+localtime)
    except:
        print("pass")
        pass
for _ in range(3):
    main()
