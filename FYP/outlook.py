import win32com.client
import pyttsx3
def main():

    outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox=outlook.GetDefaultFolder(6)
    messages=inbox.Items
    messages.Sort("[ReceivedTime]", True)
    print("Outlook notifications: ")
    n=1
    for message in messages:
        n=n+1
        if message.UnRead == True:

            print (message.Subject," from ",message.Sender) #or whatever command you want to do
            print("\n")
        if n>10:
           break
'''
message=messages.GetFirst()
print(message)
'''
if __name__ == '__main__':
    main()



