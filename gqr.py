import qrcode
direccion = input("Escribe aquí la liga:")
nombre= input("Nombre de salida:")

compuesto=nombre+".png"
img = qrcode.make(direccion)
f = open(compuesto, "wb")
img.save(f)
f.close()
print("Listo...")
