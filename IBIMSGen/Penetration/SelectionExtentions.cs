using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
using System.Collections.Generic;

namespace IBIMSGen
{
    public static class SelectionExtentions
    {
        public static List<Element> PickElements(this UIDocument uidoc, Func<Element, bool> validateElement, IPickElementsOption pickElementsOption)
        {
            return pickElementsOption.PickElements(uidoc, validateElement);
        }
        public interface IPickElementsOption
        {
            List<Element> PickElements(UIDocument uidoc, Func<Element, bool> validateElement);
        }
        public class CurrentDocumentOption : IPickElementsOption
        {
            public List<Element> PickElements(UIDocument uidoc, Func<Element, bool> validateElement)
            {
                return uidoc.Selection.PickObjects(ObjectType.Element,
                    SelectionFilterFactory.CreateElementSelectionFilter(validateElement)).
                    Select(r => uidoc.Document.GetElement(r.ElementId))
                    .ToList();
            }
        }
        public class BothDocumentOption : IPickElementsOption
        {
            public List<Element> PickElements(UIDocument uidoc, Func<Element, bool> validateElement)
            {
                var doc = uidoc.Document;
                var refrences = uidoc.Selection.PickObjects(ObjectType.PointOnElement,
                    SelectionFilterFactory.CreateLinkableSelectionFilter(doc, validateElement));
                var elements = new List<Element>();
                foreach (var refrence in refrences)
                {
                    if (doc.GetElement(refrence.ElementId) is RevitLinkInstance linkInstance)
                    {
                        var element = linkInstance.GetLinkDocument().GetElement(refrence.LinkedElementId);
                        elements.Add(element);
                    }
                    else
                    {
                        elements.Add(doc.GetElement(refrence.ElementId));
                    }
                }
                return elements;

            }
        }
        public class linkDocumentOption : IPickElementsOption
        {
            public List<Element> PickElements(UIDocument uidoc, Func<Element, bool> validateElement)
            {
                var doc = uidoc.Document;
                var refrences = uidoc.Selection.PickObjects(ObjectType.LinkedElement,
                    SelectionFilterFactory.CreateLinkableSelectionFilter(doc, validateElement));
                var elements = refrences.Select(r => (doc.GetElement(r.LinkedElementId) as RevitLinkInstance)?.GetLinkDocument()
                        .GetElement(r.ElementId)).ToList();
                return elements;

            }
        }
        /// <summary>
        /// Make a factory for easy use them without new .... ()
        /// </summary>
        public static class PickElementsOptionFactory
        {
            public static CurrentDocumentOption CreateCurrentDocumentOption() => new CurrentDocumentOption();
            public static linkDocumentOption CreateLinkDocumentOption() => new linkDocumentOption();
            public static BothDocumentOption bothDocumentOption() => new BothDocumentOption();

        }
        
    }
    public class ElementSelectionFilter : ISelectionFilter
    {
        private readonly Func<Element, bool> validateElement;
        private readonly Func<Reference, bool> validateRefrence;

        // option 1 if the user dont wanna specify the refrence just filltering the elements//
        public ElementSelectionFilter(Func<Element, bool> ValidateElement)
        {
            validateElement = ValidateElement;
        }
        // option 2 if user wanna specify the refrence//
        public ElementSelectionFilter(Func<Element, bool> ValidateElement, Func<Reference, bool> ValidateRefrence) : this(ValidateElement)
        {
            validateRefrence = ValidateRefrence;


        }
        public bool AllowElement(Element elem)
        {
            return validateElement(elem);

        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            // if the user didnt use the second delegater the validate refrence will be null so user are free to select any refrence//
            return validateRefrence?.Invoke(reference) ?? true;
        }
    }
    public class LinkableSelectionFilter : ISelectionFilter
    {
        private readonly Func<Element, bool> LvalidateElement;
        private readonly Func<Reference, bool> LvalidateRefrence;
        private readonly Document _doc;
        public LinkableSelectionFilter(Document doc
            , Func<Element, bool> ValidateElement)
        {
            LvalidateElement = ValidateElement;
            _doc = doc;

        }
        public LinkableSelectionFilter(Document doc
            , Func<Element, bool> ValidateElement, Func<Reference, bool> ValidateRefrence) : this(doc, ValidateElement)
        {
            LvalidateRefrence = ValidateRefrence;
        }
        public bool AllowElement(Element elem) => true;
        public bool AllowReference(Reference reference, XYZ position)
        {
            if (!(_doc.GetElement(reference.ElementId) is RevitLinkInstance linkInstance)) return LvalidateElement(_doc.GetElement(reference.ElementId));
            var element = linkInstance.GetLinkDocument().GetElement(reference.LinkedElementId);
            return LvalidateElement(element);
        }
    }
    public static class SelectionFilterFactory
    {
        public static LinkableSelectionFilter CreateLinkableSelectionFilter(Document doc, Func<Element, bool> validateElement)
        {
            return new LinkableSelectionFilter(doc, validateElement);
        }
        public static ElementSelectionFilter CreateElementSelectionFilter(Func<Element, bool> validateElement)
        {
            return new ElementSelectionFilter(validateElement);
        }
    }
}
