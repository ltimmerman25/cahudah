import win32com.client as win32
word = win32.gencache.EnsureDispatch('Word.Application')

print(word.__module__)
word.Visible = True
_ = input("Press ENTER to quit:")

word.Application.Quit()