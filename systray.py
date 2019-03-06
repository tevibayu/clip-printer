from infi.systray import SysTrayIcon

def say_hello(systray):
    print("Hello World!")
    
menu_options = (("List Printer", None, say_hello),)
systray = SysTrayIcon("Printer.ico", "Clip Printer", menu_options)
systray.start()
