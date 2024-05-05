// DETECTAR LA PRESION DE LAS TECLAS PARA LOS CODIGOS ESPECIFICADOS
onClipEvent(enterFrame){
	if (Key.isDown(cizq)) _root.izq = true
	else _root.izq = false
	if (Key.isDown(carr)) _root.arr = true
	else _root.arr = false
	if (Key.isDown(cder)) _root.der = true
	else _root.der = false
	if (Key.isDown(caba)) _root.aba = true
	else _root.aba = false
	if (Key.isDown(claser)) _root.laser = true
	else _root.laser = false
}
/*
w = 87
s = 83
a = 65
d = 68
*/