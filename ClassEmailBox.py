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

    def __init__(self, emailkey, subjectkey, bodykey):
        """Las caracteristicas basicas para encontrar una caja son requeridas para
        crear un objeto de la clase
        """
        self.emailkey = emailkey
        self.subjectkey = subjectkey
        self.bodykey = bodykey

    def setStart(self, startdate):
        """Este atributo sirve para saber que día llegó la caja y de ahí comenzar a buscar"""
        self.startdate = startdate

    def setCaja(self, caja):
        """Este define cual es la caja"""        
        self.caja = caja

    def getCaja(self):
        """Regresa el numero de caja"""
        return self.caja

    def setRow(self, row):
        """Este atributo sirve agregar donde se encuentra la fecha en las rows."""        
        self.row = row

    def getRow(self):
        """Regresa la Row donde se encontró la fecha."""
        return self.row

    def setEnd(self, enddate):
        """Este atributo sirve para limitar el día de busqueda."""        
        self.enddate = enddate

    def getEmailKey(self):
        """Regresa la key que se busca en el email."""        
        return self.emailkey

    def getSubjectKey(self):
        """Regresa la key que se busca en el asunto."""        
        return self.subjectkey

    def getBodyKey(self):
        """Regresa el la key que se busca en el body"""        
        return self.bodykey
            
    def getStart(self):
        """Regresa start date de donde se iterará para buscar la caja"""
        return self.startdate

    def getEnd(self):
        """Regresa End Date  de donde se terminará de iterará para buscar la caja"""        
        return self.enddate