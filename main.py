from ventanas import Ventana
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from manage_excel import create_excel_format, update_image_excel
from publish import publish


def main():
    resolucion = 'Alta Resolución'
    if resolucion == 'Alta Resolución':
        geometry = '750x670'
    else:
        geometry = '700x550'

    config = {
        'titulo': 'Estados',
        'geometry': geometry,
        'ampliacion': 0,
        'background': '#FE7C05',
        'resolucion': resolucion,
        'path_titulo': 'Image/MLTITULO.png',
        'menu': True,
        'frame': Estados,
        'guardar_en': 'RUTAESTADOS',
        'header': True,
        'canvas': None,
        'zoomed': None,
        'expand': None
    }

    app = Ventana(config)
    app.mainloop()


class Estados(Frame):
    def __init__(self, controlador):
        super().__init__(controlador, bg='white')
        self.controlador = controlador

        self.excel = None
        self.update_excel = None
        self.carpeta = None

        self.directorio_carpeta = StringVar()
        self.directorio_excel = StringVar()
        self.directorio_update_excel = StringVar()
        self.username = StringVar()
        self.password = StringVar()

        TRANSPORTEIND = Label(self, text="Correo Electrónico", bg='white', font=('Helvetica', 11))
        TRANSPORTEIND.grid(row=0, column=0, sticky=NW, padx=10, pady=5)

        PORTALIND = Label(self, text="Contraseña", bg='white', font=('Helvetica', 11))
        PORTALIND.grid(row=1, column=0, sticky=NW, padx=10, pady=5)

        self.username_entry = Entry(self, width="10", textvariable=self.username, state="normal", font=('Helvetica', 11), bd=0, bg='#D1D1D1')
        self.username_entry.grid(row=0, column=1, sticky="EW", padx=4, pady=4, columnspan=3)

        self.password_entry = Entry(self, width="10", textvariable=self.password, state="normal", font=('Helvetica', 11), show="*", bd=0, bg='#D1D1D1')
        self.password_entry.grid(row=1, column=1, sticky="EW", padx=4, pady=4, columnspan=3)

        Label(self, text='Seleccionar Excel de datos', font=('Helvetica', 11, 'bold'), bg='white').grid(row=2, column=0, sticky=W, padx=5, pady=10, columnspan=2)
        Button(self, command=self.buscar_excel, text='Seleccionar Excel', bg='#A1A1A1', bd=0, width=20, fg="#fff", font=(None, 9)).grid(row=3, column=0, padx=5, pady=10, ipady=3, ipadx=3, sticky=W)
        Entry(self, width="65", state='readonly', textvariable=self.directorio_excel, font=(None, 9), bd=0, bg='#D1D1D1').grid(row=3, column=1, padx=5, pady=10)

        Button(self, command=self.publish_products, text='Publicar Productos', bd=0, background="#FE7C05", width=20, fg="#fff", font=(None, 9)).grid(row=4, column=0, padx=5, pady=10, ipady=3, ipadx=3, sticky=W)

        ttk.Separator(self, orient=HORIZONTAL).grid(row=5, column=0, columnspan=10, sticky="EW", pady=20, padx=10, ipadx=200)

        Label(self, text='Formato Excel', font=('Helvetica', 11, 'bold'), bg='white').grid(row=6, column=0, sticky=W, padx=5, pady=10, columnspan=2)
        Button(self, text='Generar Plantilla Excel', bg='#A1A1A1', fg='white', bd=0, font=(None, 9), command=create_excel_format).grid(row=6, column=1, padx=140, pady=10, ipady=5, ipadx=5, sticky=W)

        Label(self, text='Seleccionar carpeta que contiene imagenes de los productos', font=('Helvetica', 11, 'bold'), bg='white').grid(row=8, column=0, sticky=W, padx=5, pady=10, columnspan=2)
        Label(self, text='El nombre de cada subcarpeta debe ser el código del producto respectivo. EJ: RQT-102', font=('Helvetica', 11), bg='white').grid(row=9, column=0, sticky=W, padx=5, pady=10, columnspan=2)
        Button(self, command=self.search_directory, text='Seleccionar carpeta', background="#A1A1A1", fg="#fff", width=20, bd=0, font=(None, 9)).grid(row=10, column=0, padx=5, pady=10, ipady=3, ipadx=3, sticky=W)
        Entry(self, width="65", state='readonly', textvariable=self.directorio_carpeta, font=(None, 9), bd=0, bg='#D1D1D1').grid(row=10, column=1, padx=5, pady=10)

        Label(self, text='Seleccionar Excel con los códigos de los productos', font=('Helvetica', 11, 'bold'), bg='white').grid(row=11, column=0, sticky=W, padx=5, pady=10, columnspan=2)
        Button(self, command=self.search_excel, text='Seleccionar Excel', bg='#A1A1A1', bd=0, fg="#fff", width=20, font=(None, 9)).grid(row=12, column=0, padx=5, pady=10, ipady=3, ipadx=3, sticky=W)
        Entry(self, width="65", state='readonly', textvariable=self.directorio_update_excel, font=(None, 9), bd=0, bg='#D1D1D1').grid(row=12, column=1, padx=5, pady=10)

    def search_directory(self):
        try:
            self.carpeta = filedialog.askdirectory()
            self.directorio_carpeta.set(self.carpeta)
        except AttributeError:
            pass

    def search_excel(self):
        try:
            if self.carpeta:
                self.update_excel = filedialog.askopenfile(filetypes=[("Excel files", "*.xlsx")]).name
                self.directorio_update_excel.set(self.update_excel)
                update_image_excel(self.update_excel, self.carpeta)
            else:
                messagebox.showinfo('Info', 'Falta seleccionar la carpeta con las imagenes')
        except AttributeError:
            pass

    def buscar_excel(self):
        try:
            self.excel = filedialog.askopenfile(filetypes=[("Excel files", "*.xlsx")]).name
            self.directorio_excel.set(self.excel)
        except AttributeError:
            pass

    def publish_products(self):
        if self.excel:
            if self.username.get() and self.password.get():
                counter, success = publish(self.excel, self.username.get(), self.password.get())
                messagebox.showinfo('Info', 'Se publicaron ' + str(success) + ' de ' + str(counter) + ' procuctos')
            else:
                messagebox.showinfo('Info', 'Falta ingresar las credenciales de acceso')
        else:
            messagebox.showinfo('Info', 'Falta ingresar el excel con los datos')


if __name__ == "__main__":
    main()
