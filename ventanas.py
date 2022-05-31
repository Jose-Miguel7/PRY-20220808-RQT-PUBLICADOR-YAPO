from tkinter import *
from tkinter import ttk


class Ventana(Tk):
    def __init__(self, config_ventana, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ------------------------------- AJUSTE VENTANA ---------------------------------------
        self.title(config_ventana['titulo'])
        self.geometry(config_ventana['geometry'])
        self.resizable(config_ventana['ampliacion'], config_ventana['ampliacion'])
        self.iconbitmap('./data/icon.ico')
        self.config(bg='white')
        if config_ventana['zoomed'] is not None:
            self.state('zoomed')

        # ------------------------------- ATRIBUTOS Y FRAME --------------------------------------------
        self.resolucion = config_ventana['resolucion']

        if config_ventana['canvas'] is None and config_ventana['expand'] is None:

            if config_ventana['header'] == True:
                FrameSuperior = Frame(self)
                FrameSuperior.config(bg=config_ventana['background'])
                FrameSuperior.grid(row=0, column=0, sticky='EWNS', ipadx=2000)
                title = Label(self, text="Publicador Yapo", bg=config_ventana['background'],
                              font=('Arial', 18, 'bold'), fg='#fff')
                title.grid(row=0, column=0, sticky=NW, columnspan=5, ipadx=100, ipady=25)

            FrameContenido = config_ventana['frame'](self)
            FrameContenido.grid(row=1, column=0, sticky='EWNS', pady=20, padx=30)

        else:
            if config_ventana['canvas'] == True:
                contenedor = Frame(self)
                contenedor.pack()
                self.canvas = CanvasWindow(contenedor, self, config_ventana['frame'])
            elif config_ventana['expand'] == True:
                FrameContenido = config_ventana['frame'](self)
                FrameContenido.pack(expand=True)

    """
        # ------------------------------ MENU ----------------------------------------------------------
        self.guardar_en = None
        if config_ventana['menu'] == True:
            MENU = Menu(self)
            self.config(menu=MENU)
            Guardar = Menu(MENU, tearoff=0)
            Guardar.add_command(label="Carpeta", command=self.GuardarEn)
            MENU.add_cascade(label="Guardar en", menu=Guardar)
            self.guardar_en = config_ventana['guardar_en']

    def GuardarEn(self):
        if self.guardar_en is not None:
            RutaSeleccionada = filedialog.askdirectory(parent=self)
            GuardarRuta = Ruta(self.guardar_en, RutaSeleccionada)
            GuardarRuta.guardar()
    """

    def configure_scroll_region(self):
        self.configure(scrollregion=self.bbox("all"), width=1050, height=650)


class CanvasWindow(Canvas):
    def __init__(self, contenedor, controlador, frame, *args, **kwargs):
        super().__init__(contenedor, *args, **kwargs, highlightthickness=0, bg='white')

        self.frameVenta = frame(self, controlador)
        self.frameVenta.columnconfigure(0, weight=1)
        scrollbar = ttk.Scrollbar(contenedor, orient="vertical", command=self.yview)
        self.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.pack(side="left")
        self.scrollable_window = self.create_window((0, 0), window=self.frameVenta, anchor="nw")

        def configure_scroll_region(event):
            self.configure(scrollregion=self.bbox("all"), width=1050, height=650)

        self.frameVenta.bind("<Configure>", configure_scroll_region)
