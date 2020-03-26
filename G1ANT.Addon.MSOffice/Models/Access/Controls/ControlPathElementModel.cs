using System;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class ControlPathElementModel
    {
        public const char ElementSeparator = '=';
        public const string DefaultPropertyName = "Name";

        public string Element { get; }
        public int ChildIndex { get; } = -1;
        public string PropertyName { get; } = "";
        public string PropertyValue { get; } = "";

        public ControlPathElementModel(string element)
        {
            ValidateElement(element);
            Element = element;

            var elementParts = element.Split(ElementSeparator);
            ValidateElementParts(elementParts);

            var propertyName = elementParts.Length > 1 ? elementParts[0] : DefaultPropertyName;
            var propertyValue = elementParts.Last();
            var childIndex = -1;

            if (ContainsIndex(propertyValue))
            {
                var childIndexValue = propertyValue.Substring(propertyValue.IndexOf('[') + 1);
                childIndexValue = childIndexValue.Substring(0, childIndexValue.IndexOf(']'));
                if (!int.TryParse(childIndexValue.Trim(), out childIndex) || childIndex < 0)
                    throw new ArgumentException($"Index of element ({childIndexValue}) is not a positive integer");
                propertyValue = propertyValue.Substring(0, propertyValue.IndexOf('['));
            }

            ChildIndex = childIndex;
            PropertyName = propertyName;
            PropertyValue = propertyValue;
        }

        private static void ValidateElementParts(string[] pathParts)
        {
            if (pathParts.Length > 2 || pathParts.Length == 0)
            {
                throw new ArgumentOutOfRangeException(
                    "element",
                    "Path element can contain only value for name property (like `nameOfControl`) " +
                    $"or name of property and its value separated by {ElementSeparator} (like `Caption{ElementSeparator}captionOfControl`)" +
                    $"and/or numeric child index (`[0]` or `nameOfControl[0]` or `ControlType{ElementSeparator}111[2]`)"
                );
            }
        }

        private static void ValidateElement(string element)
        {
            if (string.IsNullOrEmpty(element))
                throw new ArgumentNullException(nameof(element));
        }

        private static bool ContainsIndex(string propertyValue)
        {
            return propertyValue.Contains('[') && propertyValue.Contains(']');
        }

        public bool ShouldFilterByPropertyNameAndValue() => PropertyValue != "" || ChildIndex < 0;

        public bool ShouldFilterByIndex() => ChildIndex >= 0;
    }
}
