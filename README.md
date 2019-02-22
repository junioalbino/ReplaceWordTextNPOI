# Replace Word Text Using NPOI

## Introduction

The code bellow shows a program written in [C#](https://docs.microsoft.com/pt-br/dotnet/csharp/) with [NPOI](https://github.com/tonyqus/npoi) that reads a Word file, goes through each block, and remove paragraphs that are between some tags, eg **BEGIN** and **END** tags.

It is useful for cases when you want to make an output Word file with distinct parts for each case, eg. contracts that vary for different costumers, and you want to do this in a dynamic way.

## Background

The solution is a very simple console application, but uses some advanced concepts of C# and OOP, like [polymorphism](https://docs.microsoft.com/dotnet/csharp/programming-guide/classes-and-structs/polymorphism) and [extension methods](https://docs.microsoft.com/dotnet/csharp/programming-guide/classes-and-structs/extension-methods).

## Using the code

Create a new console application project on Visual Studio and paste the code on Program.cs. Or you can use [your favovite editor](https://code.visualstudio.com/), in this case, type the following code to create the project using [.Net Cli](https://docs.microsoft.com/pt-br/dotnet/core/tools/?tabs=netcore2x):

`dotnet new console`

Add the NPOI package to the project. To do this, on Visual Studio right-click on project and select "Manage Nuget Packages", search and install NPOI. Or use the following command on console:

`dotnet new package npoi`

Replace the Program.cs file with the following code:

```C#
namespace ReplaceWordText
{
    class Program
    {
        static void Main(string[] args)
        {
            var doc = new XWPFDocument(OPCPackage.Open("input.docx"));
            doc.RemoveParagraphs("BEGIN", "END");
            doc.Write(new FileStream("output.docx", FileMode.Create));
        }
    }

    static class SWPFDocumentExtensions
    {
        public static void RemoveParagraphs(this XWPFDocument doc, string beginTag, string endTag)
        {
            var remove = false;

            int i = 0;
            while (i < doc.BodyElements.Count)
            {
                if (!(doc.BodyElements[i] is XWPFParagraph))
                {
                    i++;
                    continue;
                }

                var runText = (doc.BodyElements[i] as XWPFParagraph).Text;

                if (runText == beginTag)
                    remove = true;

                if (remove)
                    doc.RemoveBodyElement(i);

                if (runText == endTag)
                {
                    remove = false;
                    continue;
                }

                if (!remove)
                    i++;
            }
        }
    }
}
```

## Points of Interest

I thought interesting using extension methods C#'s feature because I think it would be nice if this already was inside NPOI. But you don't have to use extension method, this solution can be done with others means.
