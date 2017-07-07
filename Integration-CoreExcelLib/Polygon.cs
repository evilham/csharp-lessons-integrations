using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_CoreExcelLib
{
    class Polygon
    {
        virtual public string Name
        {
            get { return "Das ist ein Polygon"; }
        }
		// Privates Feld, darf nur von eigenem Objekt zugegriefen werden!
        private Point2D[] points;

        // Standard Konstruktor
        public Polygon() { }
        public Polygon(Point2D[] points)
        {
			// Achtung!
			// Hier wird die *Eigenschaft* gesetzt, nicht direkt das Feld
            Points = points;
        }
		// Öffentliche Eigenschaft für die Punkte!
        public Point2D[] Points
        {
            get { return points; } // Mit 'get', geben wir nur den Array aus
            set
            {
				// Mit 'set' wird den Punkt Array zuerst ersetzt:
                points = value;
				// Und danach wird die Fläche neu berechnet
                Area = calculateArea();
            }
        }
		// Privates Feld.
        private double area;
		// Öffentliche Eigenschaft mit Privatem Setter.
		// Darf gelesen werden aber nur von eigenem Objekt gesetzt.
        public double Area
        {
            get { return area; }
            private set
            {
                area = value;
            }
        }
        private double calculateArea()
        {
            return AreaCalculator.CalculateArea(Points.ToArray());
        }
    }
    class Triangle : Polygon
    {
        public override string Name
        {
            get { return "Das ist ein Dreieck"; }
        }
        public Triangle() : base()
        {
        }
        public Triangle(Point2D[] points) : base(points)
        {
        }
    }
}
