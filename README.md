# PivotWrap
> Pivot Table Wrapper Class for Excel VBA

## About

The pivot table is arguably Excel's most powerful tool. It can summarize so much data in so little time with no more than a few clicks and drags. But creating and controlling one using VBA is a different story. The amount of boilerplate required to build the simplest of pivot tables is mind-bloggling. If you're an Excel power user who enjoy writing code from scratch to interact with pivot tables in your VBA projects, then you're probably the only one out there, end if.

PivotWrap is a set of VBA class modules that streamlines the pivot table coding experience for greater productivity. It takes care of the boilerplates for you so you can focus on building pivot tables of all complexities using concise, intuitive class methods. 

## Installing PivotWrap

Open up the Visual Basic Editor (Alt+F11) in your macro-enabled project workbook, then import (Ctrl+M) both clsPt.cls and clsPtField.cls files. Once imported, you'll see them in the Class Modules folder in the Project Explorer (Ctrl+R).

You'll also need to enable the Microsoft Scripting Runtime reference. Go to Tools/References and set a tick the checkbox beside Microsoft Scripting Runtime. This reference enables the use of the Scripting library and the Scripting.Dictionary class that is used in PivotWrap. 

## Using PivotWrap

### Initializing Class Module

First, select the Range of the source data.

```vba
Dim rng as Range
Set rng = ActiveWorkbook.Worksheets("SourceData").Range("A3:J199")
```

Make sure that the top row of the Range is the header row. As best practice, each header should be unique. Note that Excel allows you to create a pivot table with duplicate headers in the source data (It adds a number suffix to one of them in the PivotTable Fields but the source data headers remain duplicated). In contrast, PivotWrap **will not allow duplicate headers in the source data**. It will **directly overwrite** the offending header in the source data with an added suffix. 

Once the source data Range is defined, PivotWrap can be initialized.

```vba
Dim pt As PtW
Set pt = New PtW

pt.init rng
```

The ``init`` method accepts a Range argument, which it uses as the SourceData for creating a pivot cache.

### Adding Pivot Fields





