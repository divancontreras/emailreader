class BoxExcel(object):
    """Esta clase sirve para integrar los datos de la caja, es decir:
        Información de correo y datos de excel.

    Attributes:
        Email
        subject
        adsubject
        Caja: El numero de la caja.
        Sello
        TipoCaja
        Destino
        Salida
        fecha
        hora
        row

    """

    def __init__(self, email, subject, adsubject):
        """Las caracteristicas basicas para encontrar una caja son requeridas para
        crear un objeto de la clase
        """
        self.email = email
        self.subject = subject
        self.adsubject = adsubject

    def getEmail(self):
        return self.email

    def getSubject(self):
        return self.subject
        
    def getAdSubject(self):
        return self.adsubject

    def getCaja(self):
        return self.caja

    def getSello(self):
        return self.sello
        
    def getTipoCaja(self):
        return self.tipocaja

    def getDestino(self):
        return self.destino
        
    def getSalida(self):
        return self.salida
                
    def getFecha(self):
        return self.fecha

    def getHora(self):
        return self.hora

    def getRow(self):
        return self.row

    def setCaja(self, caja):
        self.caja = caja

    def setBody(self, body):
        self.body = body

    def getBody(self):
        return self.body 

    def setSello(self, sello):
        self.sello = sello
        
    def setTipoCaja(self,tipocaja):
        self.tipocaja = tipocaja

    def setDestino(self, destino):
        self.destino = destino
        
    def setSalida(self, salida):
        self.salida = salida
                
    def setFecha(self,fecha):
        self.fecha = fecha

    def setHora(self, hora):
        self.hora = hora

    def setRow(self, row):
        self.row = row
        
    def getEmail(self):
        return self.email

    def setTipoCaja(self,tipocaja):
        self.tipocaja = tipocaja
