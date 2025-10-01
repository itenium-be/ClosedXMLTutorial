using ClosedXML.Attributes;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXmlTutorial.Util;

/// <summary>
/// Where can I get one?
/// </summary>
public class SalesGenerator
{
    public IEnumerable<Sell> Generate(int amount)
    {
        yield return new Sell("Nails", 3.99M, 37);
        yield return new Sell(Name: "Hammer", Price: 12.10M, Quantity: 5, Discount: 0.1M);
        yield return new Sell("Saw", 15.37M, 12);
    }
}

public record Sell(
    string Name,
    decimal Price,
    int Quantity, 
    decimal? TotalPlaceholder = null, // Being overwritten with a formula
    decimal? Discount = null
);
