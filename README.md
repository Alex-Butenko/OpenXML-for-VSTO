# OpenXML-for-VSTO
OpenXML-for-VSTO provides the ability to process content as OpenXML from VSTO-addins

* [Introduction](#introduction)
* [How it works?](#How-it-works)
* [Issues](#Issues)
* [Where to use](#Where-to-se)
* [Where not to use](#Where-not-to-use)

## Introduction
One of the problems of creating VSTO add-ins for MS Office is the performance of access to a large number of elements. Access is through the boundaries of the AppDomains, with a lot of Reflection. Each call can read/write only one parameter at a time, therefore, for large tables/documents, the add-in have to make millions of such calls. Moreover, these calls are performed in the UI thread, blocking or slowing down user input and worsening the user experience.
Using OpenXML would be a good solution, but unfortunately, the Office API does not provide easy access to OpenXML.
This library provides the ability to get/set an OpenXML object directly from/to VSTO add-in.

## Where to use
In VSTO add-ins, in operations with thousands of objects, involving styles and formats.

## Where not to use
* Office Interop except for VSTO add-ins. In this case the library does not give any advantage, as you can replace interop with any OpenXML library all at once.
* Operations with small sets of objects - each copy paste has significant overhead that does not worth it for small sets.
* Operations with values/formulas in Excel - these can be much faster accessed throuch 2d arrays.

## Benchmark
### Excel
Benchmark is based on writing and reading cells - 1, 100, 10,000 or 1,000,000 two different ways - using this library (and ClosedXML) and with pure Excel interop. Each individual cell value, background color, font color, font size, italic and bold status and number format. No optimizations.

Writing operations:

|                 | Pure VSTO  | Using OpenXML-for-VSTO |
|-----------------|------------|------------------------|
| 1 cell          |00:00:00.254|    00:00:00.938        |
| 100 cells       |00:00:00.302|    00:00:01.225        |
| 10000 cells     |00:00:09.709|    00:00:01.967        |
| 1,000,000 cells |00:18:31.897|    00:01:28.600        |

Reading operations:

|                 | Pure VSTO  | Using OpenXML-for-VSTO |
|-----------------|------------|------------------------|
| 1 cell          |00:00:00.009|    00:00:00.754        |
| 100 cells       |00:00:00.045|    00:00:00.740        |
| 10000 cells     |00:00:03.287|    00:00:01.476        |
| 1,000,000 cells |00:05:45.179|    00:00:47.838        |

Even though performance for Excel can be improved in some scenarios by using union ranges, ClosedXML algorithms are less than optimal in this scenario as well, and most likely in real scenario performance gain can be much bigger.

### Word
Benchmark is based on writing and reading runs of text - 1, 100, 10,000, 100,000 or 1,000,000 two different ways - using this library  and with pure Word interop. Each individual run text, background color, font color, font size, italic and bold status. No optimizations.

Writing operations:

|                 | Pure VSTO  | Using OpenXML-for-VSTO |
|-----------------|------------|------------------------|
| 1 run           |00:00:00.027|     00:00:00.828       |
| 100 runs        |00:00:00.623|     00:00:00.583       |
| 10,000 runs     |00:02:10.232|     00:00:01.360       |
| 100,000 runs    |00:51:37.007|     00:00:09.061       |
| 1,000,000 runs  |            |     b00:01:36.512      |

Reading operations:

|                 | Pure VSTO  | Using OpenXML-for-VSTO |
|-----------------|------------|------------------------|
| 1 run           |00:00:00.013|     00:00:00.519       |
| 100 runs        |00:00:00.150|     00:00:00.186       |
| 10,000 runs     |00:00:16.284|     00:00:01.377       |
| 100,000 runs    |00:03:02.905|     00:00:08.515       |
| 1,000,000 runs  |            |     00:01:23.920       |


## How it works
* Copy an object to a file using Clipboard.
* Work with the file as with any other OpenXML file.
* Copy the object back.

## Issues
* It uses the clipboard, so after each operation it looses data user potentiall saved there. There are workaronds, like save and restore clipboard for each operation, but these are not part of this library
* Beginning from certain size (around 12 mb) of objects, the library may cause IsolatedStorageException. This exception happens because VSTO appdomain has incorrect Evidence. It is easy to fix by calling this code once before using OpenXML-for-VSTO with big objects:
```c#
System.Security.Policy.Evidence newEvidence = new System.Security.Policy.Evidence();
newEvidence.AddHostEvidence(new System.Security.Policy.Zone(System.Security.SecurityZone.MyComputer));

System.AppDomain.CurrentDomain
    .GetType()
    .GetField("_SecurityIdentity", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
    .SetValue(System.AppDomain.CurrentDomain, newEvidence);
```
