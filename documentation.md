# Presentation

## Overview of PowerPoint Presentation

Essential Presentation is a native .NET class library that can be used by developers to create, read, and
write Microsoft PowerPoint files by using C#, VB.NET, and managed C++ code. The library can be used in
Windows Forms, WPF, ASP.NET, ASP.NET MVC, UWP and Xamarin platforms.

It is a non-UI component that provides a full-fledged PowerPoint presentation instance that facilitates
accessing and manipulating the presentations without any dependency of Microsoft Office COM
libraries and Microsoft Office.

### Key features

*   Support to create PowerPoint presentation from scratch.
*   Open, modify, and save existing presentations.
*   Ability to convert PowerPoint presentation to PDF.
*   Ability to convert PowerPoint slides to images.
*   Ability to create and edit charts.
*   Ability to convert chart in a slide to image.
*   Ability to clone and merge slides in presentation
*   Ability to create and edit animations.
*   Ability to create and edit transition effects.
*   Ability to create and edit comments.
*   Ability to encrypt and decrypt PowerPoint presentation.
*   Ability to set and remove write protection of PowerPoint presentation.
*   Ability to access the Built-in and Custom document properties.
*   Ability to create and modify sections in PowerPoint presentation.

### Compatible Microsoft PowerPoint Versions

*   Microsoft PowerPoint 2007
*   Microsoft PowerPoint 2010
*   Microsoft PowerPoint 2013
*   Microsoft PowerPoint 2016
*   Microsoft PowerPoint 2019

**Note:**
1.  The current version of Essential Presentation supports the .PPTX, .PPTM, .POTX, .POTM file formats only.
2.  The current version of Essential Presentation does not support some features in Microsoft
    PowerPoint such as Word Art, creation and editing of Handouts, equations, create and edit audio and
    video content, built-in themes, and its variants.

### Assemblies Required

The following assemblies need to be referenced in your application

| Platform(s)                                 | Assembly                                                                                         |
| ------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| WPF, Windows Forms, ASP. NET and ASP.NET MVC | Syncfusion.Presentation.Base<br>Syncfusion.Compression.Base<br>Syncfusion.OfficeChart.Base         |
| ASP.NET Core, Xamarin and Blazor            | Syncfusion.Presentation.Portable<br>Syncfusion.Compression.Portable<br>Syncfusion.OfficeChart.Portable |
| Universal Windows Platform                  | Syncfusion.Presentation.UWP<br>Syncfusion.OfficeChart.UWP                                       |

**Note:** Starting with v16.2.0.x, if you reference Syncfusion assemblies from trial setup or from the NuGet
feed, you also have to include a license key in your projects. Please refer to this link to know about
registering Syncfusion license key in your applications to use our components.

### Converting PowerPoint Presentation to PDF

For converting a PowerPoint Presentation to PDF, the following assemblies needed to be referenced in
your application

| Platform(s)                                 | Assembly                                                                                                                                                             |
| ------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| WPF, Windows Forms, ASP. NET and ASP.NET MVC | Syncfusion.Presentation.Base<br>Syncfusion.Compression.Base<br>Syncfusion.OfficeChart.Base<br>Syncfusion.Pdf.Base<br>Syncfusion.PresentationToPDFConverter.Base         |
| ASP.NET Core, Xamarin and Blazor            | Syncfusion.Presentation.Portable<br>Syncfusion.Compression.Portable<br>Syncfusion.OfficeChart.Portable<br>Syncfusion.Pdf.Portable<br>Syncfusion.PresentationRenderer.Portable<br>Syncfusion.SkiaSharpHelper.Portable<br>Skiasharp |

The following assemblies are required to be referred in addition to the above mentioned assemblies for
converting the chart present in the PowerPoint Presentation into PDF.

| Platform(s)                                 | Assembly                                               |
| ------------------------------------------- | ------------------------------------------------------ |
| WPF, Windows Forms, ASP. NET and ASP.NET MVC | Syncfusion.OfficeChartToImageConverter.WPF<br>Syncfusion.SfChart.WPF |

**Note:** 1.The “Syncfusion.OfficeChartToImageConverter.WPF” assembly is supported from .NET
Framework 4.0 onwards

### NuGet Packages Required

To work with PowerPoint Presentations, install the following NuGet packages in your application:

| Platform(s)                                                        | NuGet Package                               |
| ------------------------------------------------------------------ | ------------------------------------------- |
| Windows Forms, Console Application (Targeting .NET Framework)      | Syncfusion.Presentation.WinForms.nupkg      |
| WPF                                                                | Syncfusion.Presentation.Wpf.nupkg           |
| .NET Framework 3.5 or 4.0 Client Profile                           | Syncfusion.Presentation.ClientProfile.nupkg |
| ASP.NET Web Forms                                                  | Syncfusion.Presentation.AspNet.nupkg        |
| ASP.NET MVC4                                                       | Syncfusion.Presentation.AspNet.Mvc4.nupkg   |
| ASP.NET MVC5                                                       | Syncfusion.Presentation.AspNet.Mvc5.nupkg   |
| UWP                                                                | Syncfusion.Presentation.UWP.nupkg           |
| ASP.NET Core, Console Application (Targeting .NET Core) and Blazor | Syncfusion.Presentation.Net.Core.nupkg      |
| Xamarin                                                            | Syncfusion.Xamarin.Presentation.nupkg       |

**Note:** 1.Starting with v16.2.0.x, if you reference Syncfusion assemblies from trial setup or from the
NuGet feed, add the "Syncfusion.Licensing" assembly reference and include a license key in your
projects. Refer to this link to learn about registering Syncfusion license key in your applications to use
the components.
2.From the Essential Studio 2018 Volume 3 release(v16.3.0.21), Syncfusion has changed some of the
NuGet package names to search and find the required Syncfusion NuGet packages in nuget.org easily
based on the control and its platforms.

### Getting Started

#### Creating a simple PowerPoint Presentation with basic elements from scratch

In this page, you can learn how to create a simple PowerPoint Presentation by using Essential
Presentation API.
For creating and manipulating a PowerPoint Presentation, include the following assemblies in the
application.

| Assembly Name              | Short Description                                                                          |
| -------------------------- | ------------------------------------------------------------------------------------------ |
| Syncfusion.Presentation.Base | This assembly contains the core features required for creating, reading, manipulating a Presentation file. |
| Syncfusion.Compression.Base  | This assembly is used to package the Presentation contents.                                |
| Syncfusion.OfficeChart.Base  | This assembly contains the office chart object model and core features needed for chart creation. |

**Note:** Starting with v16.2.0.x, if you reference Syncfusion assemblies from trial setup or from the NuGet
feed, you also have to include a license key in your projects. Please refer to this link to know about
registering Syncfusion license key in your applications to use our components.

Include the following namespace in your .cs or .vb code as shown below

**C#**
```csharp
using Syncfusion.Presentation;
```

**VB.NET**
```vbnet
Imports Syncfusion.Presentation
```

**UWP**
```csharp
using Syncfusion.Presentation;
```

**ASP.NET CORE**
```csharp
using Syncfusion.Presentation;
```

**XAMARIN**
```csharp
using Syncfusion.Presentation;
```

An entire PowerPoint Presentation is represented by an instance of 'IPresentation' interface and it is the
root element of Essential Presentation’s DOM.
The following code example demonstrates how to create an instance of 'IPresentation' interface.

**C#**
```csharp
//Creates a new instance of PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
```

**VB.NET**
```vbnet
'Creates a new instance of PowerPoint presentation
Dim pptxDoc As IPresentation = Presentation.Create()
```

**UWP**
```csharp
//Creates a new instance of PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
```

**ASP.NET CORE**
```csharp
//Creates a new instance of PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
```

**XAMARIN**
```csharp
//Creates a new instance of PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
```

'IPresentation' instance has a slide collection that represents the individual slides present within
PowerPoint presentation. A slide may contain textual and other graphics contents like shapes, images,
charts etc.
The following code example demonstrates how to add a blank slide to a PowerPoint Presentation.

**C#**
```csharp
//Adds a slide to the PowerPoint Presentation
ISlide firstSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
```

**VB.NET**
```vbnet
'Adds a slide to the PowerPoint Presentation
Dim firstSlide As ISlide = pptxDoc.Slides.Add(SlideLayoutType.Blank)
```

**UWP**
```csharp
//Adds a slide to the PowerPoint Presentation
ISlide firstSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
```

**ASP.NET CORE**
```csharp
//Adds a slide to the PowerPoint Presentation
ISlide firstSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
```

**XAMARIN**
```csharp
//Adds a slide to the PowerPoint Presentation
ISlide firstSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
```

**Note:** The 'Point' typographic units are used to add or manipulate any element in a Presentation.

All the textual contents in a Presentation document are represented by paragraphs. Within the
paragraph, textual contents are grouped into one or more child elements as 'TextParts'. Each 'TextPart'
represents a region of text with a common set of formatted text.
The following code example demonstrates how to add text into a presentation.

**C#**
```csharp
//Adds a textbox in a slide by specifying its position and size
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);
//Adds a paragraph into the textShape
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Set the horizontal alignment of paragraph
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
//Adds a textPart in the paragraph
ITextPart textPart = paragraph.AddTextPart("Hello Presentation");
//Applies font formatting to the text
textPart.Font.FontSize = 80;
textPart.Font.Bold = true;
```

**VB.NET**
```vbnet
'Adds a textbox in a slide by specifying its position and size
Dim textShape As IShape  = firstSlide.AddTextBox(100, 75, 756, 200)
'Adds a paragraph into the textShape
Dim paragraph As IParagraph  = textShape.TextBody.AddParagraph()
'Set the horizontal alignment of paragraph
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center
'Add a textPart in the paragraph
Dim textPart As ITextPart  = paragraph.AddTextPart("Hello Presentation")
'Applies font formatting to the text
textPart.Font.FontSize = 80
textPart.Font.Bold = True
```

**UWP**
```csharp
//Adds a textbox in a slide by specifying its position and size
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);
//Adds a paragraph into the textShape
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Set the horizontal alignment of paragraph
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
//Adds a textPart in the paragraph
ITextPart textPart = paragraph.AddTextPart("Hello Presentation");
//Applies font formatting to the text
textPart.Font.FontSize = 80;
textPart.Font.Bold = true;
```

**ASP.NET CORE**
```csharp
//Adds a textbox in a slide by specifying its position and size
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);
//Adds a paragraph into the textShape
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Set the horizontal alignment of paragraph
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
//Adds a textPart in the paragraph
ITextPart textPart = paragraph.AddTextPart("Hello Presentation");
//Applies font formatting to the text
textPart.Font.FontSize = 80;
textPart.Font.Bold = true;
```

**XAMARIN**
```csharp
//Adds a textbox in a slide by specifying its position and size
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);
//Adds a paragraph into the textShape
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Set the horizontal alignment of paragraph
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
//Adds a textPart in the paragraph
ITextPart textPart = paragraph.AddTextPart("Hello Presentation");
//Applies font formatting to the text
textPart.Font.FontSize = 80;
textPart.Font.Bold = true;
```

Essential Presentation allows you to create simple and multi-level lists that make the content easier for
reading. The following code example demonstrates how to add a bulleted list in a paragraph.

**C#**
```csharp
//Adds a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the
fictitious company on which the AdventureWorks sample databases are based,
is a large, multinational manufacturing company.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the font of the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the hanging value as 20
paragraph.FirstLineIndent = -20;
```

**VB.NET**
```vbnet
'Adds a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the
fictitious company on which the AdventureWorks sample databases are based,
is a large, multinational manufacturing company.")
'Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted
'Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183)
'Sets the font of the bullet character
paragraph.ListFormat.FontName = "Symbol"
'Sets the hanging value as 20
paragraph.FirstLineIndent = -20
```

**UWP**
```csharp
//Adds a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the
fictitious company on which the AdventureWorks sample databases are based,
is a large, multinational manufacturing company.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the font of the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the hanging value as 20
paragraph.FirstLineIndent = -20;
```

**ASP.NET CORE**
```csharp
//Adds a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the
fictitious company on which the AdventureWorks sample databases are based,
is a large, multinational manufacturing company.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the font of the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the hanging value as 20
paragraph.FirstLineIndent = -20;
```

**XAMARIN**
```csharp
//Adds a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the
fictitious company on which the AdventureWorks sample databases are based,
is a large, multinational manufacturing company.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the font of the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the hanging value as 20
paragraph.FirstLineIndent = -20;
```

In PowerPoint Presentation, the multilevel lists are used for presenting the content in a hierarchy. You
can create a multi-level list by setting the indentation levels. By default, the level begins at 0 and
increments by 1 for each level. The following code example demonstrates how to add multi-level list in a
paragraph.

**C#**
```csharp
//Adds a new paragraph
paragraph = textShape.TextBody.AddParagraph("The company manufactures and
sells metal and composite bicycles to North American, European and Asian
commercial markets.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the list level as 2. Possible values can range from 0 to 8
paragraph.IndentLevelNumber = 2;
```

**VB.NET**
```vbnet
'Adds a new paragraph
paragraph = textShape.TextBody.AddParagraph("The company manufactures and
sells metal and composite bicycles to North American, European and Asian
commercial markets.")
'Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted
'Sets the list level as 2. Possible values can range from 0 to 8
paragraph.IndentLevelNumber = 2
```

**UWP**
```csharp
//Adds a new paragraph
paragraph = textShape.TextBody.AddParagraph("The company manufactures and
sells metal and composite bicycles to North American, European and Asian
commercial markets.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the list level as 2. Possible values can range from 0 to 8
paragraph.IndentLevelNumber = 2;
```

**ASP.NET CORE**
```csharp
//Adds a new paragraph
paragraph = textShape.TextBody.AddParagraph("The company manufactures and
sells metal and composite bicycles to North American, European and Asian
commercial markets.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the list level as 2. Possible values can range from 0 to 8
paragraph.IndentLevelNumber = 2;
```

**XAMARIN**
```csharp
//Adds a new paragraph
paragraph = textShape.TextBody.AddParagraph("The company manufactures and
sells metal and composite bicycles to North American, European and Asian
commercial markets.");
//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the list level as 2. Possible values can range from 0 to 8
paragraph.IndentLevelNumber = 2;
```

You can add images to the Presentation by adding them in the picture collection of a slide. The following
code example demonstrates how to add an image in a presentation.

**C#**
```csharp
//Gets the image from file path
Image image = Image.FromFile(@"image.jpg");
// Adds the image to the slide by specifying position and size
firstSlide.Pictures.AddPicture(new MemoryStream(image.ImageData), 300, 270,
410, 250);
```

**VB.NET**
```vbnet
'Gets the image from file path
Dim image__1 As Image = Image.FromFile("image.jpg")
' Adds the image to the slide by specifying position and size
firstSlide.Pictures.AddPicture(New MemoryStream (image__1.ImageData), 300,
270, 410, 250)
```

**UWP**
```csharp
//Gets the image from file path
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream imageStream =
assembly.GetManifestResourceStream("UWP.Data.tablet.jpg");
// Adds the image to the slide by specifying position and size
firstSlide.Pictures.AddPicture(imageStream, 300, 270, 410, 250);
```

**ASP.NET CORE**
```csharp
//Gets the image from file path
FileStream imageStream = new FileStream(@"Image.png", FileMode.Open,
FileAccess.Read);
// Adds the image to the slide by specifying position and size
firstSlide.Pictures.AddPicture(imageStream, 300, 270, 410, 250);
```

**XAMARIN**
```csharp
//Gets the image from file path
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream imageStream =
assembly.GetManifestResourceStream("SampleBrowser.Presentation.Samples.Template.tablet.jpg");
// Adds the image to the slide by specifying position and size
firstSlide.Pictures.AddPicture(imageStream, 300, 270, 410, 250);
```

Finally, save the Presentation in file system and close its instance.

**C#**
```csharp
//Saves the Presentation in the given name
pptxDoc.Save("Output.pptx");
//Releases the resources occupied
pptxDoc.Close();
```

**VB.NET**
```vbnet
'Saves the Presentation in the given name
pptxDoc.Save("Output.pptx")
'Releases the resources occupied
pptxDoc.Close()
```

**UWP**
```csharp
//Initializes FileSavePicker
FileSavePicker savePicker = new FileSavePicker();
savePicker.SuggestedStartLocation = PickerLocationId.Desktop;
savePicker.SuggestedFileName = "Sample";
savePicker.FileTypeChoices.Add("PowerPoint Files", new List<string>() {
".pptx" });
//Creates a storage file from FileSavePicker
StorageFile storageFile = await savePicker.PickSaveFileAsync();
//Saves changes to the specified storage file
await pptxDoc.SaveAsync(storageFile);
//Releases the resources occupied
pptxDoc.Close();
```

**ASP.NET CORE**
```csharp
//Saving the PowerPoint Presentation as stream
FileStream stream = new FileStream("Sample.pptx", FileMode.Create,
FileAccess.ReadWrite);
pptxDoc.Save(stream);
//Dispose stream
stream.Dispose();
//Close the presentation
pptxDoc.Close();
```

**XAMARIN**
```csharp
//Create new memory stream to save Presentation.
MemoryStream stream = new MemoryStream();
//Save Presentation in stream format.
pptxDoc.Save(stream);
//Close the presentation
pptxDoc.Close();
stream.Position = 0;
if (Device.OS == TargetPlatform.WinPhone || Device.OS ==
TargetPlatform.Windows)
Xamarin.Forms.DependencyService.Get<ISaveWindowsPhone>().Save("Sample.pptx",
"application/vnd.openxmlformats-officedocument.presentationml.presentation",
stream);
else
Xamarin.Forms.DependencyService.Get<ISave>().Save("Sample.pptx",
"application/vnd.openxmlformats-officedocument.presentationml.presentation",
stream);

## Document Object Model representation

In order to create and modify a PowerPoint Presentation, you need to know how the elements are organized in Essential Presentation’s document object model (DOM). The following figure illustrates this DOM.

*(DOM Image from PDF page 2030)*

## Load and save the Presentation

### Opening an existing Presentation from file system
You can open an existing PowerPoint Presentation by using the file name and its physical path.
**C#**
```csharp
//Opens an existing Presentation from file system
IPresentation pptxDoc = Presentation.Open(fileName);
```
**VB.NET**
```vbnet
'Opens an existing Presentation from file system
Dim pptxDoc As IPresentation = Presentation.Open(fileName)
```
**UWP**
```csharp
//Instantiates the File Picker
FileOpenPicker openPicker = new FileOpenPicker();
openPicker.SuggestedStartLocation = PickerLocationId.Desktop;
openPicker.FileTypeFilter.Add(".pptx");
//Creates a storage file from FileOpenPicker
StorageFile inputStorageFile = await openPicker.PickSingleFileAsync();
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = await Presentation.OpenAsync(inputStorageFile);
```
**ASP.NET CORE**
```csharp
//Loads or open an PowerPoint Presentation
FileStream inputStream = new FileStream("Sample.pptx", FileMode.Open);
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(inputStream);
```
**XAMARIN**
```csharp
//"App" is the class of Portable project
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream inputStream = assembly.GetManifestResourceStream("Sample.pptx");
//Loads or open an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(inputStream);
```

### Opening an existing Presentation from stream
You can open an existing PowerPoint Presentation from stream by using the overloads of Open method.
**C#**
```csharp
//Opens an existing Presentation from stream
IPresentation pptxDoc = Presentation.Open(presentationStream);
```
**VB.NET**
```vbnet
'Opens an existing Presentation from stream
Dim pptxDoc As IPresentation = Presentation.Open(presentationStream)
```
**UWP**
```csharp
//Create new Presentation without slides.
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream inputStream = assembly.GetManifestResourceStream(inputFilePath);
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(inputStream);
```
**ASP.NET CORE**
```csharp
//Loads or open an PowerPoint Presentation
FileStream inputStream = new FileStream(inputFileName, FileMode.Open);
```
**XAMARIN**
```csharp
//Create new Presentation without slides.
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream inputStream = assembly.GetManifestResourceStream(inputFilePath);
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(inputStream);
```

### Opening an encrypted Presentation
You can open an encrypted PowerPoint presentation from either file path or stream by using the following overloads of Open method as follows.
**C#**
```csharp
//Opens an existing encrypted Presentation from stream
IPresentation pptxDoc = Presentation.Open(presentationStream, password);
//Opens an existing encrypted Presentation from file system
IPresentation pptxDoc = Presentation.Open(fileName, password);
```
**VB.NET**
```vbnet
'Opens an existing encrypted Presentation from stream
Dim pptxDoc As IPresentation = Presentation.Open(presentationStream, password)
'Opens an existing encrypted Presentation from file system
Dim pptxDoc As IPresentation = Presentation.Open(fileName, password)
```
**UWP**
```csharp
//Opens an existing encrypted Presentation from stream
IPresentation pptxDoc = Presentation.OpenAsync(presentationStream, password);
//Opens an existing encrypted Presentation from file system
IPresentation pptxDoc = Presentation.OpenAsync(fileName, password);
```
**Note:** Essential Presentation Library does not provides support to Encryption and Decryption in ASP.NET Core and Xamarin platforms.

### Saving a PowerPoint Presentation to file system
You can save the created or manipulated PowerPoint Presentation to file system by using `Save()` method of `IPresentation` interface. Default format type is `*.PPTX`.
**C#**
```csharp
//Opens an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(fileName);
//To-Do some manipulation
//Saves the Presentation in file system
pptxDoc.Save("Output.pptx");
```
**VB.NET**
```vbnet
'Opens an existing PowerPoint Presentation
Dim pptxDoc As IPresentation = Presentation.Open(fileName)
'To-Do some manipulation
'Saves the Presentation in file system
pptxDoc.Save("Output.pptx")
```
**UWP**
```csharp
//Instantiates the File Picker
FileOpenPicker openPicker = new FileOpenPicker();
openPicker.SuggestedStartLocation = PickerLocationId.Desktop;
openPicker.FileTypeFilter.Add(".pptx");
//Creates a storage file from FileOpenPicker
StorageFile inputStorageFile = await openPicker.PickSingleFileAsync();
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = await Presentation.OpenAsync(inputStorageFile);
//To-Do some manipulation
//Initializes FileSavePicker
FileSavePicker savePicker = new FileSavePicker();
savePicker.SuggestedStartLocation = PickerLocationId.Desktop;
savePicker.SuggestedFileName = "Sample";
savePicker.FileTypeChoices.Add("PowerPoint Files", new List<string>() { ".pptx" });
//Creates a storage file from FileSavePicker
StorageFile storageFile = await savePicker.PickSaveFileAsync();
//Saves changes to the specified storage file
await pptxDoc.SaveAsync(storageFile);
```
**ASP.NET CORE**
```csharp
//Loads or open an PowerPoint Presentation
FileStream inputStream = new FileStream(fileName, FileMode.Open);
//To-Do some manipulation
FileStream outputStream = new FileStream("output.pptx", FileMode.Create);
pptxDoc.SaveAs(outputStream);
```
**XAMARIN**
```csharp
//"App" is the class of Portable project
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream inputStream = assembly.GetManifestResourceStream(inputFilePath);
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(inputStream);
//To-Do some manipulation
//Create new memory stream to save Presentation.
MemoryStream stream = new MemoryStream();
//Save Presentation in stream format.
pptxDoc.Save(stream);
//Close the presentation
pptxDoc.Close();
stream.Position = 0;
//The operation in Save under Xamarin varies between Windows Phone, Android and iOS platforms.
if (Device.OS == TargetPlatform.WinPhone || Device.OS == TargetPlatform.Windows)
    Xamarin.Forms.DependencyService.Get<ISaveWindowsPhone>().Save("Output.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", stream);
else
    Xamarin.Forms.DependencyService.Get<ISave>().Save("Output.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", stream);
```

### Saving a PowerPoint Presentation to stream
You can save the created or manipulated PowerPoint Presentation to stream by using overloads of `Save` method.
**C#**
```csharp
//Opens an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(fileName);
//To-Do some manipulation
//Creates an instance of memory stream
MemoryStream stream = new MemoryStream();
//Saves the Presentation to stream
pptxDoc.Save(stream);
```
**VB.NET**
```vbnet
'Opens an existing PowerPoint Presentation
Dim pptxDoc As IPresentation = Presentation.Open(fileName)
'To-Do some manipulation
'Creates an instance of memory stream
Dim stream As New MemoryStream()
'Saves the Presentation to stream
pptxDoc.Save(stream)
```
**UWP**
```csharp
//Instantiates the File Picker
FileOpenPicker openPicker = new FileOpenPicker();
openPicker.SuggestedStartLocation = PickerLocationId.Desktop;
openPicker.FileTypeFilter.Add(".pptx");
//Creates a storage file from FileOpenPicker
StorageFile inputStorageFile = await openPicker.PickSingleFileAsync();
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = await Presentation.OpenAsync(inputStorageFile);
//To-Do some manipulation
//Saves changes to the specified storage file
MemoryStream outputStream = new MemoryStream();
await pptxDoc.SaveAsync(outputStream);
```
**ASP.NET CORE**
```csharp
//Loads or open an PowerPoint Presentation
FileStream inputStream = new FileStream(inputFileName, FileMode.Open);
//To-Do some manipulation
FileStream outputStream = new FileStream(outputFileName, FileMode.Create);
pptxDoc.SaveAs(outputStream);
```
**XAMARIN**
```csharp
//"App" is the class of Portable project
Assembly assembly = typeof(App).GetTypeInfo().Assembly;
Stream inputStream = assembly.GetManifestResourceStream(inputFilePath);
//Loads or open an PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(inputStream);
//To-Do some manipulation
MemoryStream outputStream = new MemoryStream();
pptxDoc.Save(outputStream);
```

### Sending to a client browser
You can save and send the Presentation to a client browser from a website or web application by invoking the overload of `Save` method. This method explicitly make use of an instance of `HttpResponse` as its parameter in order to stream the presentation to client browser. So, this overload is suitable for web application that refer to `System.Web` assembly.
**C#**
```csharp
//Opens an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(fileName);
//To-Do some manipulation
//Saves the Presentation to the client browser
pptxDoc.Save("Output.pptx", FormatType.Pptx, Response);
```
**VB.NET**
```vbnet
'Opens an existing PowerPoint Presentation
Dim pptxDoc As IPresentation = Presentation.Open(fileName)
'To-Do some manipulation
'Saves the Presentation to the client browser
pptxDoc.Save("Output.pptx", FormatType.Pptx, Response)
```
**ASP.NET CORE**
```csharp
//Loads or open an PowerPoint Presentation
FileStream inputStream = new FileStream(inputFileName, FileMode.Open);
//To-Do some manipulation
//Initialize content type
string ContentType = null;
//Save the PowerPoint Presentation to stream
MemoryStream outputStream = new MemoryStream();
pptxDoc.Save(outputStream);
outputStream.Position = 0;
//Return the file with content type
return File(outputStream, ContentType, outputFileName);
```
**Note:** Saving and sending the workbook to a client browser from a web site is suitable for web applications alone.

### Closing a PowerPoint Presentation
When you are done with the Presentation instance, you should close the instance of `IPresentation` in order to release the memory consumed by Essential Presentation library.
**C#**
```csharp
//Opens an existing Presentation from file system
IPresentation pptxDoc = Presentation.Open(fileName);
//To-Do some manipulation
//Creates an instance of memory stream
MemoryStream stream = new MemoryStream();
//Saves the Presentation to stream
pptxDoc.Save(stream);
//Closes the Presentation instance and free the memory consumed.
pptxDoc.Close();
```
**VB.NET**
```vbnet
'Opens an existing Presentation from file system
Dim pptxDoc As IPresentation = Presentation.Open(fileName)
'To-Do some manipulation
'Creates an instance of memory stream
Dim stream As New MemoryStream()
'Saves the Presentation to stream
pptxDoc.Save(stream)
'Closes the Presentation instance and free the memory consumed.
pptxDoc.Close()
```

## Working with PowerPoint presentation

### Cloning a PowerPoint presentation
Cloning a PowerPoint presentation creates a new copy of the PowerPoint presentation and the changes made in the cloned copy of the presentation do not affect the source PowerPoint presentation.
**C#**
```csharp
//Opens a PowerPoint presentation
IPresentation sourcePresentation = Presentation.Open(fileName);
//Clones the Presentation
IPresentation clonedPresentation = sourcePresentation.Clone();
//Gets the first slide from the cloned PowerPoint presentation
ISlide firstSlide = clonedPresentation.Slides[0];
//Adds a textbox in a slide by specifying its position and size
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);
//Adds a paragraph in the body of the textShape
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Adds a textPart in the paragraph
ITextPart textPart = paragraph.AddTextPart("Essential Presentation");
//Saves the modified cloned PowerPoint presentation
clonedPresentation.Save("ClonedPresentation.pptx");
```

### Printing a PowerPoint presentation
You can print the Presentation document by converting the PowerPoint presentation slides to images. For more information about converting the PowerPoint presentation slides to images, see Conversion. You can use the `System.Drawing.Printing.PrintDocument` class to print the converted images by the default printer or to any of the available printer with customized settings.

### Working with PowerPoint presentation properties
Document properties, also known as meta data, are details about a file that describe or identify it. Document properties are classified into two categories.
*   **Built-in Document Properties** - that include details such as title, author name, subject, and keywords that identify the document's topic or contents.
*   **Custom Document properties** - define the user-defined document properties.

**Built-in Document Properties**
You can access and modify the built in document properties of a PowerPoint presentation with Essential Presentation library. The Built-in document properties of a PowerPoint presentation is represented by `IBuiltInDocumentProperties` type.

**Accessing and Modifying Built-in Document Properties**
**C#**
```csharp
//Opens a PowerPoint presentation
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Accesses the built-in document properties
Console.WriteLine("Title - {0}", pptxDoc.BuiltInDocumentProperties.Title);
Console.WriteLine("Author - {0}", pptxDoc.BuiltInDocumentProperties.Author);
//Modifies the Built-in document properties
pptxDoc.BuiltInDocumentProperties.Category = "Sales reports";
pptxDoc.BuiltInDocumentProperties.Company = "Northwind traders";
//Saves the modified PowerPoint presentation
pptxDoc.Save("Output.pptx");
//Closes the modified PowerPoint presentation
pptxDoc.Close();
```

**Custom Document properties**
You can create and modify the custom document properties of a PowerPoint presentation with Essential Presentation library. The collection of custom document properties in a PowerPoint presentation is represented by `ICustomDocumentProperties` object.

**Adding Custom Document properties**
**C#**
```csharp
//Creates a PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Adds custom document properties
ICustomDocumentProperties documentProperty = pptxDoc.CustomDocumentProperties;
documentProperty.Add("PropertyA");
documentProperty["PropertyA"].Text = "@!123";
documentProperty.Add("PropertyB");
documentProperty["PropertyB"].Text = "B";
//Saves the PowerPoint presentation
pptxDoc.Save("Output.pptx");
//Closes the PowerPoint presentation
pptxDoc.Close();
```

**Accessing and Modifying Custom Document Properties**
**C#**
```csharp
//Opens a PowerPoint presentation
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Accesses an existing custom document property
IDocumentProperty property = pptxDoc.CustomDocumentProperties["PropertyA"];
//Modifies the value of DocumentProperty
property.Value = "Hello world";
//Saves the PowerPoint presentation
pptxDoc.Save("Output.pptx");
//Closes the PowerPoint presentation
pptxDoc.Close();
```

### Marking a PowerPoint presentation as final
PowerPoint presentation can be made read-only to prevent the readers from making inadvertent changes to it. However, making presentation as final is not a security feature. Anyone can disable the final status and edit the presentation.
**C#**
```csharp
//Create an instance for PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Add slide to the presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Mark the presentation as final
pptxDoc.Final = true;
//Save the presentation
pptxDoc.Save("MarkAsFinal.pptx");
//Close the presentation
pptxDoc.Close();
```

## Working with Slides in PowerPoint

### Adding slide to the PowerPoint presentation
In PowerPoint presentation, a slide is a container for the elements like shapes, images, charts, text box etc. The slides may inherit the formatting and layout properties from its 'Master' and 'Layout' slides.
**C#**
```csharp
//Creates a PowerPoint instance
IPresentation pptxDoc = Presentation.Create();
//Adds a slide to the PowerPoint presentation
ISlide slide = pptxDoc.Slides.Add();
//Saves the Presentation to the file system.
pptxDoc.Save("Sample.pptx");
//Closes the Presentation instance
pptxDoc.Close();
```

### Create a slide with predefined LayoutSlide
The Syncfusion PowerPoint library supports the following predefined slide layout types to create a slide as equivalent to Microsoft PowerPoint:
*   Blank
*   Comparison
*   Content with caption
*   Picture with caption
*   Section header
*   Title
*   Title and content
*   Title and vertical text
*   Title only
*   Two content
*   Vertical title and text

**C#**
```csharp
//Create a PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Add a slide of blank layout type
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Save the PowerPoint file
pptxDoc.Save("Sample.pptx");
//Close the PowerPoint instance
pptxDoc.Close();
```

### Adding Custom layout slide
The slide layout are template design for the PowerPoint slides. Slide layout can contains formatting, positioning, and placeholders for a slide.
**C#**
```csharp
//Open the template presentation
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Add a new custom layout slide to the master collection with a specific layout type and name
ILayoutSlide layoutSlide = pptxDoc.Masters[0].LayoutSlides.Add(SlideLayoutType.Blank, "CustomLayout");
//Set background of the layout slide
layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(78, 89, 90);
//Get the stream of an image
Stream pictureStream = File.Open("Image.png", FileMode.Open);
//Add the picture into layout slide
layoutSlide.Shapes.AddPicture(pictureStream, 100, 100, 100, 100);
//Add a slide of new designed custom layout to the presentation
ISlide slide = pptxDoc.Slides.Add(layoutSlide);
//Save the presentation
pptxDoc.Save("Output.pptx");
//Close the presentation
pptxDoc.Close();
```

### Cloning slide
You can create a deep copy of a slide by cloning the slide. The cloned slide is an independent copy of its source slide. This means the changes made in the cloned slide do not affect the source slide.
**C#**
```csharp
//Opens an existing Presentation.
IPresentation pptxDoc = Presentation.Open("Presentation.pptx");
//Retrieves the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Creates a cloned copy of slide.
ISlide slideClone = slide.Clone();
//Adds a new text box to the cloned slide.
IShape textboxShape = slideClone.AddTextBox(0, 0, 250, 250);
//Adds a paragraph with text content to the shape.
textboxShape.TextBody.AddParagraph("Hello Presentation");
//Adds the slide to the Presentation.
pptxDoc.Slides.Add(slideClone);
//Saves the Presentation to the file system.
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Merging slide
The Essential Presentation provides ability to clone slides from one Presentation to another Presentation. With this ability, you can split a large Presentation into small ones and also merge multiple presentations to one Presentation. You can choose the theme for the cloned slide by using the enum `PasteOption`.
**C#**
```csharp
//Opens the source Presentation
IPresentation sourcePresentation = Presentation.Open("SourcePresentation.pptx");
//Opens the destination Presentation
IPresentation destinationPresentation = Presentation.Open("DestinationPresentation.pptx");
//Clones the first slide of the source Presentation
ISlide clonedSlide = sourcePresentation.Slides[0].Clone();
//Merges the cloned slide to the destination Presentation with paste option - Destination Theme
destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme, sourcePresentation);
//Saves the destination Presentation
destinationPresentation.Save("Output.pptx");
//Closes the source presentation
sourcePresentation.Close();
//Closes the destination Presentation
destinationPresentation.Close();
```

### Removing slide
The Essential Presentation provides the ability to delete a slide by its instance or by its index position in slide collection.
**C#**
```csharp
//Opens an existing presentation.
IPresentation pptxDoc = Presentation.Open("Presentation1.pptx");
//Retrieves the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Removes the specified slide from the Presentation.
pptxDoc.Slides.Remove(slide);
// Removes the slide from the specified index.
pptxDoc.Slides.RemoveAt(1);
//Saves the destination Presentation
pptxDoc.Save("Output.pptx");
//Closes the Presentation instance
pptxDoc.Close();
```

### Converting to image
You can convert a presentation slide to image with Essential Presentation.
**C#**
```csharp
//Opens a PowerPoint presentation file
IPresentation pptxDoc = Presentation.Open(fileName);
//Creates an instance of ChartToImageConverter and assigns it to ChartToImageConverter
pptxDoc.ChartToImageConverter = new ChartToImageConverter();
//Converts the first slide into image
Image image = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageType.Metafile);
//Saves the image as file
image.Save("slide1.png");
//Closes the Presentation instance
pptxDoc.Close();
```

### Changing Slide background
**C#**
```csharp
//Opens an existing Presentation.
IPresentation pptxDoc = Presentation.Open("Presentation1.pptx");
//Retrieves the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Retrieves the background instance.
IBackground background = slide.Background;
//Sets the fill type of the background to gradient.
background.Fill.FillType = FillType.Gradient;
//Retrieves the fill of the background to the IGradientFill instance.
IGradientFill gradient = background.Fill.GradientFill;
//Adds the first gradient stop of the gradient fill.
gradient.GradientStops.Add(ColorObject.Green, 20);
//Adds the second gradient stop of the gradient fill.
gradient.GradientStops.Add(ColorObject.Yellow, 50);
//Saves the Presentation to the file system
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Create and edit Master and Layout slides
To get all the slides in same format, you should perform those changes in the Slide Master or Layout Master. The changes will be applied to all the slides, which inherits the master slide or layout slide.

### Access the MasterSlide
In PowerPoint presentation, the MasterSlide is the top slide that controls all information about the theme, layout, background, color, fonts, and positioning of all slides.
**C#**
```csharp
//Create a PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Access the first master slide in PowerPoint file
IMasterSlide masterSlide = pptxDoc.Masters[0];
//Get the first shape name from the master slide
string shapeName = masterSlide.Shapes[0].ShapeName;
//Save the PowerPoint file
pptxDoc.Save("Sample.pptx");
//Close the Presentation instance
pptxDoc.Close();
```

### Create a custom LayoutSlide
The Syncfusion PowerPoint library lets you build your own custom layout designs and use them to create individual slides.
**C#**
```csharp
//Create a PowerPoint instance
IPresentation pptxDoc = Presentation.Create();
//Add a new LayoutSlide to the PowerPoint file
ILayoutSlide layoutSlide = pptxDoc.Masters[0].LayoutSlides.Add(SlideLayoutType.Blank, "CustomLayout");
//Add a shape to the LayoutSlide
IShape shape = layoutSlide.Shapes.AddShape(AutoShapeType.Diamond, 30, 20, 400, 300);
//Change the background color for LayoutSlide
layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(78, 89, 90);
//Save the PowerPoint file
pptxDoc.Save("LayoutSlide.pptx");
//Close the Presentation instance
pptxDoc.Close();
```

## Working with Paragraph

### Applying Paragraph formatting
Each paragraph in a slide can has its own formatting types such as alignment, indent etc.
**C#**
```csharp
//Loads the PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Gets the slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Gets the shape in slide
IShape textboxShape = slide.Shapes[0] as IShape;
//Gets instance of a paragraph in a textbox
IParagraph paragraph = textboxShape.TextBody.Paragraphs[0];
//Applies the first line indent of the paragraph
paragraph.FirstLineIndent = 10;
//Applies the horizontal alignment of the paragraph to center.
paragraph.HorizontalAlignment = HorizontalAlignmentType.Left;
//Applies the left indent of the paragraph
paragraph.LeftIndent = 8;
//Saves the Presentation
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Modifying text
You can modify a text by accessing the existing paragraphs in a Presentation.
**C#**
```csharp
//Opens an existing Presentation from file system.
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Retrieves the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Retrieves the first paragraph of the shape.
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Retrieves the first TextPart of the shape.
ITextPart textPart = paragraph.TextParts[0];
//Modifies the text content of the TextPart.
textPart.Text = "Hello Presentation";
//Saves the presentation to the file system.
pptxDoc.Save("Result.pptx");
//Closes the Presentation.
pptxDoc.Close();
```

### Enabling shrink text on overflow option
By using a Shrink text on overflow option, you can fit a large text within a shape.
**C#**
```csharp
// Create a new PowerPoint file.
using (IPresentation ppDoc = Presentation.Create())
{
    // Add a slide to the PowerPoint file.
    ISlide slide = ppDoc.Slides.Add(SlideLayoutType.Blank);
    // Add a text box to the slide
    IShape textBox = slide.Shapes.AddTextBox(100, 100, 100, 100);
    //Add text to the text box.
    textBox.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
    //Set the property to shrink text on overflow.
    textBox.TextBody.FitTextOption = FitTextOption.ShrinkTextOnOverFlow;
    // Save the PowerPoint file
    ppDoc.Save("Sample.pptx");
}
```
**Note:** The shrink text on overflow is not supported in UWP, ASP.NET CORE and Xamarin platforms.

### Removing the paragraph
**C#**
```csharp
//Opens an existing Presentation from file system.
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Retrieves the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape
IShape shape = slide.Shapes[0] as IShape;
//Retrieves the first paragraph of the shape
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Removes the first paragraph from the textbody of the shape
shape.TextBody.Paragraphs.Remove(paragraph);
//Saves the presentation to the file system
pptxDoc.Save("Result.pptx");
//Closes the Presentation
pptxDoc.Close();
```

## Working with lists
Essential Presentation allows you to create simple and multi-level lists that make the content easier for reading. In PowerPoint, Presentation lists consists of the following types
1.  Numbered list
2.  Bulleted list
3.  Picture list

### Numbered List
**C#**
```csharp
//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
// Adds a textbox to hold the list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 270);
// Adds a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as Numbered
paragraph.ListFormat.Type = ListType.Numbered;
//Sets the numbered style (list numbering) as Arabic number following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Sets the starting value as 1
paragraph.ListFormat.StartValue = 1;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
// Sets the hanging value
paragraph.FirstLineIndent = -20;
// Sets the bullet character size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100;
//Saves the Presentation to the file system.
pptxDoc.Save("Sample.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Bulleted list
**C#**
```csharp
//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds the slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
// Adds a textbox to hold the list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 250);
// Adds a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as bulleted
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
// Sets the font for the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Saves the Presentation to the file system.
pptxDoc.Save("Sample.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Picture List
**C#**
```csharp
//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds the slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
// Adds a textbox to hold the list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 270);
// Adds a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as Numbered
paragraph.ListFormat.Type = ListType.Picture;
//Sets the image for the list.
paragraph.ListFormat.Picture(new MemoryStream(Syncfusion.Drawing.Image.FromFile("Image.png").ImageData));
// Sets the picture size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 150;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
// Sets the hanging value
paragraph.FirstLineIndent = -20;
//Saves the Presentation to the file system.
pptxDoc.Save("Sample.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Creating a Multilevel List
You can create a multi-level list by setting the indentation levels. By default, the level begins at 0 and increments by 1 for each level. A list can be incremented or decremented from levels 0 to 8 as like Microsoft PowerPoint.
**C#**
```csharp
//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds the slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds a textbox to hold the bulleted list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 250);
//Adds paragraph to the textbox
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as Numbered list
paragraph.ListFormat.Type = ListType.Numbered;
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
paragraph.IndentLevelNumber = 1;
//Adds paragraph to the textbox
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Sets the list type as Numbered list
paragraph.ListFormat.Type = ListType.Numbered;
paragraph.ListFormat.NumberStyle = NumberedListStyle.AlphaLcPeriod;
//Sets the list level as 2
paragraph.IndentLevelNumber = 2;
//Saves the Presentation to the file system.
pptxDoc.Save("MultiLevelList.pptx");
//Closes the Presentation
pptxDoc.Close();
```

## Working with Shapes

### Adding shapes to a slide
In every slide, there is a shape collection that can contain any form of graphical objects such as AutoShape, chart, text, or picture. You can add any shape element to this collection. The `IShape` is the base type for the shape elements.
**C#**
```csharp
//Creates an instance for PowerPoint
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds normal shape to slide
slide.Shapes.AddShape(AutoShapeType.Cube, 50, 200, 300, 300);
//Creates an instance for image as stream
Stream imageStream = File.Open("Image.jpg", FileMode.Open);
//Add picture to the shape collection
IPicture picture = slide.Shapes.AddPicture(imageStream, 373, 83, 526, 382);
//Saves the Presentation
pptxDoc.Save("Sample.pptx");
//Closes the stream
imageStream.Close();
//Closes the Presentation
pptxDoc.Close();
```

### Iterating through shapes
You can iterate through the shapes in a PowerPoint slide.
**C#**
```csharp
//Opens an existing Presentation from the file system
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Iterates through shapes in a slide and sets title
foreach(IShape shape in pptxDoc.Slides[0].Shapes)
{
    if (shape is IPicture)
        shape.Title = "Picture";
    else if (shape is IShape)
        shape.Title = "AutoShape";
}
//Saves the Presentation
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Specifying shape properties
The shape properties can be used to format and modify the shapes in a slide.
**C#**
```csharp
//Creates instance for PowerPoint
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Gets the first slide of the Presentation
ISlide slide = pptxDoc.Slides[0];
//Gets the shape of the slide
IShape shape = slide.Shapes[0] as IShape;
//Sets the shape name.
shape.ShapeName = "Shape1";
//Retrieves the line format of the shape.
ILineFormat lineFormat = shape.LineFormat;
//Sets the dash style of the line format.
lineFormat.DashStyle = LineDashStyle.DashDotDot;
//Sets the weight of the line format.
lineFormat.Weight = 3;
//Sets the pattern fill type to shape
shape.Fill.FillType = FillType.Pattern;
//Chooses the type of pattern
shape.Fill.PatternFill.Pattern = PatternFillType.DashedDownwardDiagonal;
//Sets the fore color
shape.Fill.PatternFill.ForeColor = ColorObject.AliceBlue;
//Sets the back color
shape.Fill.PatternFill.BackColor = ColorObject.DarkSalmon;
//Saves the Presentation
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Removing the shapes
The shapes can be removed from a slide by its instance or by its index position in the shape collection.
**C#**
```csharp
//Opens an existing Presentation from file system
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Retrieves the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Removes the shape from the shape collection.
slide.Shapes.Remove(shape);
//Saves the Presentation to the file system.
pptxDoc.Save("Result.pptx");
//Closes the Presentation.
pptxDoc.Close();
```

### Working with GroupShape
The shapes in a slide can be grouped into a single shape.
**C#**
```csharp
//Creates an instance for PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds a group shape to the slide
IGroupShape groupShape = slide.GroupShapes.AddGroupShape(20, 20, 450, 300);
//Adds a TextBox to the group shape
groupShape.Shapes.AddTextBox(30, 25, 100, 100).TextBody.AddParagraph("My TextBox");
//Gets the image stream
Stream pictureStream = File.Open("Image.png", FileMode.Open);
//Adds a picture to the group shape
groupShape.Shapes.AddPicture(pictureStream, 40, 100, 100, 100);
//Adds a shape to the group shape
groupShape.Shapes.AddShape(AutoShapeType.Rectangle, 200, 200, 90, 30);
//Save the presentation
pptxDoc.Save("Output.pptx");
//Close the presentation
pptxDoc.Close();
```

## Working with images

### Replacing Images
**C#**
```csharp
//Opens an existing Presentation.
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Retrieves the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first picture from the slide.
IPicture picture = slide.Pictures[0];
//Gets the new picture as stream.
Stream pictureStream = File.Open("Image.png", FileMode.Open);
//Creates instance for memory stream
MemoryStream memoryStream = new MemoryStream();
//Copies stream to memoryStream.
pictureStream.CopyTo(memoryStream);
//Replaces the existing image with new image.
picture.ImageData = memoryStream.ToArray();
//Saves the Presentation to the file system.
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Removing Images
**C#**
```csharp
//Opens an existing Presentation from file system.
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Retrieves the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Iterates through the pictures collection and remove the picture
foreach (IPicture picture in slide.Pictures)
{
    //Removes the picture from the slide.
    slide.Pictures.Remove(picture);
    break;
}
//Saves the Presentation to the file system.
pptxDoc.Save("Output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

## Working with PowerPoint Tables

### Create a table by adding rows
**C#**
```csharp
//Create a PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Add slide to the presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a table to the slide
ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);
//Initialize index values to add text to table cells
int rowIndex = 0, colIndex;
//Iterate row-wise cells and add text to it
foreach (IRow rows in table.Rows)
{
    colIndex = 0;
    foreach (ICell cell in rows.Cells)
    {
        cell.TextBody.AddParagraph("(" + rowIndex.ToString() + " , " + colIndex.ToString() + ")");
        colIndex++;
    }
    rowIndex++;
}
//Save the presentation
pptxDoc.Save("Sample.pptx");
//Close the presentation
pptxDoc.Close();
```

### Applying table formatting
You can format a table to change its appearance by customizing the table border, cell background, cell margins etc.
**C#**
```csharp
//Creates instance of PowerPoint Presentation
IPresentation pptxDoc = Presentation.Create();
//Adds slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds table to the slide
ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);
//Retrieves each cell and fills text content to the cell.
ICell cell = table[0, 0];
//Sets the column width for a cell; this sets the width for entire column
cell.ColumnWidth = 400;
//Sets the margin for the cell.
cell.TextBody.MarginBottom = 0;
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Sets the back color for the cell.
cell.Fill.SolidFill.Color.SystemColor = Color.Orange;
cell.TextBody.AddParagraph("First Row and First Column");
//Saves the Presentation
pptxDoc.Save("Table.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Applying table styles
You can format a table by applying pre-defined table styles.
**C#**
```csharp
//Creates instance of PowerPoint Presentation
IPresentation pptxDoc = Presentation.Create();
//Adds slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds table to the slide
ITable table = slide.Shapes.AddTable(3, 3, 100, 120, 300, 200);
table.BuiltInStyle = BuiltInTableStyle.ThemedStyle2Accent4;
table.HasBandedRows = false;
table.HasHeaderRow = false;
table.HasBandedColumns = true;
table.HasFirstColumn = true;
table.HasLastColumn = true;
table.HasTotalRow = true;
//Saves the Presentation
pptxDoc.Save("Table.pptx");
//Closes the Presentation
pptxDoc.Close();
```

## Working with Charts

### Creating a Chart from scratch
An instance of `IOfficeChart` can be used to create or modify the charts in PowerPoint Presentation.
**C#**
```csharp
//Creates a Presentation instance
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds chart to the slide with position and size
IPresentationChart chart = slide.Charts.AddChart(100, 10, 700, 500);
//Specifies the chart title
chart.ChartTitle = "Sales Analysis";
//Sets chart data
chart.ChartData.SetValue(1, 2, "Jan");
chart.ChartData.SetValue(2, 1, 2010);
chart.ChartData.SetValue(2, 2, 60);
//Creates a new chart series with the name
IOfficeChartSerie seriesJan = chart.Series.Add("Jan");
//Sets the data range of chart series
seriesJan.Values = chart.ChartData[2, 2, 4, 2];
//Sets the data range of the category axis
chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 4, 1];
//Specifies the chart type
chart.ChartType = OfficeChartType.Column_Clustered;
//Saves the Presentation
pptxDoc.Save("sample.pptx");
//Closes the Presentation
pptxDoc.Close();
```

### Creating charts from excel sheet
You can also create a chart with the data from an existing excel worksheet.
**C#**
```csharp
//Creates a Presentation instance
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Gets the excel file as stream
MemoryStream excelStream = new MemoryStream(File.ReadAllBytes("Book1.xlsx"));
//Adds a chart to the slide with a data range from excel worksheet
IPresentationChart chart = slide.Charts.AddChart(excelStream, 1, "A1:D4", new RectangleF(100, 10, 700, 500));
//Saves the Presentation
pptxDoc.Save("output.pptx");
//Closes the Presentation
pptxDoc.Close();
```

## Working with Animations
Animations are visual effects for the objects in PowerPoint presentation. Animation effects can be grouped into four categories:
1.  Entrance
2.  Emphasis
3.  Exit
4.  Motion paths

### Adding animation effect to shapes
**C#**
```csharp
//Create an instance for PowerPoint
using (IPresentation pptxDoc = Presentation.Create())
{
    //Add a blank slide to Presentation
    ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
    //Add normal shape to slide
    IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 50, 200, 300, 300);
    //Access the animation sequence to create effects
    ISequence sequence = slide.Timeline.MainSequence;
    //Add bounce effect to the shape
    IEffect bounceEffect = sequence.AddEffect(cubeShape, EffectType.Bounce, EffectSubtype.None, EffectTriggerType.OnClick);
    //Save the Presentation
    pptxDoc.Save("Sample.pptx");
}
```

## Add and edit transitions in PowerPoint slides
Slide transitions are the motion effects that occur when you move from one slide to the next during a slide show presentation.

### Set a transition effect to a PowerPoint slide
**C#**
```csharp
//Create a PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Set the transition effect type
slide.SlideTransition.TransitionEffect = TransitionEffect.Checkerboard;
//Set the transition effect options
slide.SlideTransition.TransitionEffectOption = TransitionEffectOption.Across;
//Save the presentation
pptxDoc.Save("Sample.pptx");
//Close the presentation
pptxDoc.Close();
```

## Working with Comments

### Adding a comment
**C#**
```csharp
//Create a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Create();
//Add a slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a comment to the slide
slide.Comments.Add(10, 10, "Author1", "A1", "Can we change the font size to 20?", DateTime.Now);
//Save the Presentation
pptxDoc.Save("Comment.pptx");
//Close the Presentation
pptxDoc.Close();
```

### Replying to a comment
**C#**
```csharp
//Create a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open("Sample.pptx");
//Get the slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Get the comment in the slide
IComment comment = slide.Comments[0] as IComment;
//Add reply to the comment
slide.Comments.Add("Author2", "A2", "Yes, we can we change the font size to 20", DateTime.Now, comment);
//Save the presentation
pptxDoc.Save("ReplyComment.pptx");
//Close the Presentation
pptxDoc.Close();
```

## Working with Sections

### Creating a section
**C#**
```csharp
//Creates a PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Adds a section to the PowerPoint presentation
ISection section = pptxDoc.Sections.Add();
//Sets a name to the created section
section.Name = "SectionDemo";
//Adds a slide to the created section
ISlide slide = section.AddSlide(SlideLayoutType.Blank);
//Adds a text box to the slide
slide.AddTextBox(10, 10, 100, 100).TextBody.AddParagraph("Slide in SectionDemo");
//Saves the PowerPoint presentation
pptxDoc.Save("Section.pptx");
```

## Security

### Encrypting with password
You can protect a PowerPoint Presentation by encrypting the document by using a password.
**C#**
```csharp
//Creates an instance for Presentation
IPresentation presentation = Presentation.Create();
//Adds slide to Presentation
ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
//Adds textbox to slide
IShape shape = slide.Shapes.AddTextBox(100, 30, 200, 300);
//Adds a paragraph with text content.
IParagraph paragraph = shape.TextBody.AddParagraph("Password Protected.");
//Protects the file with password
presentation.Encrypt("PASSWORD!@1#$");
//Saves the Presentation
presentation.Save("Sample.pptx");
//Closes the Presentation
presentation.Close();
```

### Decrypting the PowerPoint Presentation
Essential Presentation provides ability to remove the encryption from the PowerPoint Presentation. You can decrypt a PowerPoint Presentation by opening it with the password.
**C#**
```csharp
//Opens an existing Presentation from file system and it can be decrypted by using the provided password.
IPresentation presentation = Presentation.Open("Sample.pptx", "PASSWORD!@1#$");
//Decrypts the document
presentation.RemoveEncryption();
//Saves the presentation
presentation.Save("Output.pptx");
//Closes the Presentation
presentation.Close();
```

### Write Protection
You can set write protection for a PowerPoint Presentation and remove protection from the write protected PowerPoint presentation.
**C#**
```csharp
//Create a new instance for PowerPoint presentation
IPresentation pptxDoc = Presentation.Create();
//Add the blank slide to the presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Set the write protection for presentation instance
pptxDoc.SetWriteProtection("MYPASSWORD");
//Saves the modified cloned PowerPoint presentation
pptxDoc.Save("Sample.pptx");
//Close the presentation instance
pptxDoc.Close();
```

## FAQ’s
1.  **Why I get an exception when trying to load a PPT file?**
    The current version of Presentation library supports only .PPTX format - Microsoft Office 2007 and later version.
2.  **Is it possible to print the Presentation slides?**
    Yes, you can print the PowerPoint presentations by using its ability to convert the slides as images and by using the `PrintDocument` class. For more details, refer to Printing
3.  **Does adding audio and video to a Presentation is supported?**
    At present, there is no support to add audio and video to Presentation by using Essential Presentation library.
4.  **What measure does Essential Presentation use to add slide elements such as textbox, shape, picture and charts?**
    We use Points to add any slide elements in a Presentation.
5.  **Does Essential Presentation supports cloning a slide in the Presentation?**
    Yes, Essential Presentation library supports cloning as follows:
    *   Slide in the Presentation can be cloned from one Presentation to another or within a same Presentation.
    *   An entire Presentation can also be cloned as an independent copy of the original.
6.  **Could not find Syncfusion.OfficeChartToImageConverter assembly in .NET 3.5 Framework, does it mean there is no support for chart conversion in this framework?**
    Yes, OfficeChartToImageConverter assembly is not supported in .NET 3.5 Framework and it is available from .NET 4.0 Framework.
7.  **Can chart data be refreshed?**
    Yes, Essential Presentation supports refreshing the chart data. For more details, refer to Working with charts
8.  **Is it possible to convert 3D charts to PDF or image?**
    Current version of the Essential Presentation library does not provide support for converting 3D charts to PDF or image format.
9.  **How to improve the image quality while converting the Presentation slides to image?**
    You can improve the quality of converted images by specifying the image resolution. Refer – Converting PowerPoint presentation to Images
