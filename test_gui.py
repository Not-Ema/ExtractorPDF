import traceback
try:
    import run
    print('Run import OK')
    from src.gui import ExtractorApp
    print('GUI class OK')
    import tkinter as tk
    print('Tkinter OK')
    root = tk.Tk()
    app = ExtractorApp(root)
    print('App created successfully')
    root.after(500, root.destroy)
    root.mainloop()
    print('Test completed successfully')
except Exception as e:
    print('Error:', str(e))
    traceback.print_exc()