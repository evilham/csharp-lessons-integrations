using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_CoreExcelLib
{
    class Point2D
    {
        public double X = 0.0; // Default Wert
        public double Y = 0.0; // Default Wert
		/* Alle Klassen Erben von der "Object" Klasse.
		Deswegen existiert schon ein leerer Konstruktor ohne Argumente:
		public Point2D() {
			// macht nichts
		}
		Da wir X und Y schon als 0.0 initialisieren, brauchen wir den
		argumentlosen Konstruktor nicht definieren
		*/
		
		// Da die Argumente und die Felder im Objekt anders heissen,
		// braucht man kein "this.X".
		// Wäre die Signatur des Konstruktors:
		// public Point2D(double X, double Y)
		// Müsste er so aussehen:
		// public Point2D(double X, double Y) {
		//	this.X = X;
		//	this.Y = Y;
		// }
        public Point2D(double x, double y)
        {
            X = x;
            Y = y;
        }
		// Abstand von _diesem_ Punkt zu p2
        public double distanceTo(Point2D p2)
        {
			// Wenn diese Methode aufgerufen wird existiert unser Objekt schon
			// deswegen kann "this" benutzt werden.
			// Bei der Kompilierung weiss der Kompiler, dass mit "this" eine
			// Objekt Instanz diser Klasse gemeint ist:
            return distance(this, p2);
        }
		// Das ist eine statische Methode, hier können Instanz
		// Felder / Eigenschaften / Methoden nicht benutzt werden.
		// z.B. "this" ist innerhalb der Methode nicht definiert
		// (wird vom Kompiler nicht zugelassen).
        public static double distance(Point2D p1, Point2D p2)
        {
			// Einfach: Wurzel((x2 - x1)^2 + (y2 - y1)^2)
            return Math.Sqrt(
                Math.Pow(p2.X - p1.X, 2) +
                Math.Pow(p2.Y - p1.Y, 2));
        }

    }
}
