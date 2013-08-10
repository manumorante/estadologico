// FUNCIONES GLOBALES
// EASING
MovieClip.prototype.ea = function(mc, x, y) {
	mc._x += (x-mc._x)/6;
	mc._y += (y-mc._y)/6;
};