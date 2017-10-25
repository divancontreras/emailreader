class EmailBox(object):
    """Esta clase sirve para integrar los datos de la caja, es decir:
        Información de correo y datos de excel.

    Attributes:
        Caja: El numero de la caja.
        EmailKey: El key que se busca en el email.
        SubjectKey: El key que se busca en el asunto.
        BodyKey: El key que se busca en el body.
        Start: Fecha donde se inicia busqueda.
        End: Fecha donde concluye busqueda
        Row: Fila donde se encuentra la fecha.

    """

    def __init__(self, emailkey, datecol, hourcol, iteration = False):
        """Las caracteristicas basicas para encontrar una caja son requeridas para
        crear un objeto de la clase
        """
        self.emailkey = emailkey
        self.datecol = datecol
        self.hourcol = hourcol
        self.iteration = iteration

    def getDateColumn(self):
        """Este define cual es la caja"""        
        return self.datecol

    def getIteration(self):
        """Este define cual es la caja"""        
        return self.iteration

    def getHourColumn(self):
        """Este define cual es la caja"""        
        return self.hourcol

    def getFactura(self):
        """Este define cual es la caja"""        
        return self.caja

    def setFecha(self, justdate):
        """Este atributo sirve para saber que día llegó la caja y de ahí comenzar a buscar"""
        self.justdate = justdate

    def setFactura(self, caja):
        """Este define cual es la caja"""        
        self.caja = caja

    def getFactura(self):
        """Este define cual es la caja"""        
        return self.caja

    def setRow(self, row):
        """Este atributo sirve agregar donde se encuentra la fecha en las rows."""        
        self.row = row

    def getRow(self):
        """Regresa la Row donde se encontró la fecha."""
        return self.row

    def getEmailKey(self):
        """Regresa la key que se busca en el email."""        
        return self.emailkey

    def getFecha(self):
        """Regresa start date de donde se iterará para buscar la caja"""
        return self.justdate
