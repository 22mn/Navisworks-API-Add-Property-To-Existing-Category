using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace Lab8_Exercise_V2
{
    [PluginAttribute("Lab8_V2", "TwentyTwo", DisplayName = "Lab8_Exec_V2", 
        ToolTip = "A tutorial project by TwentyTwo")]

    public class MainClass : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            // current document
            Document doc = Application.ActiveDocument;
            // COM state object
            InwOpState10 cdoc = ComApiBridge.State;
            // current selected items
            ModelItemCollection items = doc.CurrentSelection.SelectedItems;
            // convert ModelItem to COM Path
            InwOaPath citem = ComApiBridge.ToInwOaPath(items[0]);
            // Get item's PropertyCategoryCollection object
            InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);
            // Get PropertyCategoryCollection data
            InwGUIAttributesColl propCol = cpropcates.GUIAttributes();
            // loop propertycategory
            foreach (InwGUIAttribute2 i in propCol)
            {                
                // if category's name match
                if ( i.UserDefined && i.ClassUserName == "Premier League")
                {
                    // overwritten the existing propertycategory with 
                    // newly created propertycategory(existing + new) 
                    cpropcates.SetUserDefined(1, "Premier League", "PremierLeague_InternalName", 
                        AddNewPropertyToExtgCategory(i));
                }
                              
            }

            return 0;
        }

        // add new property to existing category
        public InwOaPropertyVec AddNewPropertyToExtgCategory(InwGUIAttribute2 propertyCategory)
        {
            // COM state (document)
            InwOpState10 cdoc = ComApiBridge.State;
            // a new propertycategory object
            InwOaPropertyVec category = (InwOaPropertyVec)cdoc.ObjectFactory(
                nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

            // retrieve existing propertydata (name & value) and add to category
            foreach (InwOaProperty property in propertyCategory.Properties())
            {
                // create a new Property (PropertyData)
                InwOaProperty extgProp = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, 
                    null, null);
                // set PropertyName
                extgProp.name = property.name;
                // set PropertyDisplayName
                extgProp.UserName = property.UserName;
                // set PropertyValue
                extgProp.value = property.value;
                // add to category
                category.Properties().Add(extgProp);
            }

            // create a new PropertyData and add to category
            InwOaProperty newProp = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty,
                null, null);
            newProp.name = "2021Champions_Internal";
            newProp.UserName = "2021 Champions";
            newProp.value = "Who Knows!!";
            category.Properties().Add(newProp);
            return category;
        }
    }
}


