Project:	FormShaper ActiveX DLL
----------------------------------------
By:		Andrew Davey
Email:		andrewdavey@hotmail.com
----------------------------------------

(Enable word wrap to read this file.)

About:
	FormShaper lets you change the shape that windows draws your forms. You can choose from Rectangles, RoundedRectangles, Ellipses, Polygons and any combination of the above.

Install:
	Extract the ZIP file to folder. Open "ShaperTest.vbg" and click "Make FormShaper.dll" on the File menu (make sure the dll project is selected first). This should have registered the DLL for you. Or you can do it manually using "regsvr32.exe".
	Play with the example, look at the code, etc.

Use:
	Add a reference to the DLL in your project. Create an instance of CShaper. Call the add... methods, e.g. myShaper.addRect(0, 0, 100, 100). Note addPolygon() requires a zero based, array of type PointAPI. Then call myShaper.ShapeForm Me.hWnd . This will shape the form - easy!
	To save coding a lot of complex shapes you can save a shape description in a file. In a test form put calls the add... methods, then call SaveShape(strFileName). This will save your shape. To load this shape on any other form call LoadShape(strFileName).
	Finally, you can say how windows combines the shapes. By default it just overlays them all, but it can do other things. Set the optional parameter "CombineMode" of the add... method to do this. Try RGN_XOR for that swiss cheese effect!

Legal Stuff:
	This DLL is provided to you for free. Do not sell it. You may use it in your own projects providing that you ask me first (by email). I accept no responsibilty for any loss or damage to your system or data due to the use of this software. Use at your own risk. Have a nice day :-)