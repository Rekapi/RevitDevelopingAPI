// 001- Retrieving any element parameter by its name and datatype 
  // Select any element the plugin will work when you select an element first unless will give u an error
 var element = uiDocument.Selection.GetElementIds().Select(x => doc.GetElement(x)).First();

 // Retriving any parameter of the element by its name and type of input
 var value = element.LookupParameter("Area").AsValueString();

  // Retriving any parameter of the element by using BuiltInParamter
  var value = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsValueString();

 // 002- Converting Wall to solid
            //Select Wall
            // from outside class called ElementSelectionFilter
            ElementSelectionFilter wallSelection = new ElementSelectionFilter();
            // IList<Reference> refernce = uiDocument.Selection.PickObjects(ObjectType.Element, wallSelection);
            Wall SelectedWall = doc.GetElement(uiDocument.Selection.PickObject(ObjectType.Element, wallSelection)) as Wall;

            // converting wall to Solid to get its area and volume
            GeometryElement wallGeometry = SelectedWall.get_Geometry(new Options());
            Solid wallSolid = null;
            foreach (var geomObject in wallGeometry)
            {
                if(geomObject is Solid)
                {
                    Solid solid = geomObject as Solid;
                    wallSolid = solid;
                }
            }

