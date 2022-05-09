# rowlistlib2

General purpose list and row objects
Superseeds the [rowlistlib](https://github.com/francescofoti/rowlistlib).

Two main additions come with this new versions of the rowlistlib project:
 - Parsing and outputting JSON, which is done by the integration and adaptation (not easy btw) of the VBA-JSON (v2.2.2) class (renamed here CJsonConverter) from [Tim Hall (C)](https://github.com/VBA-tools/VBA-JSON);
 - CList and CRow Object type support in CList/CRow objects, which makes it possible to have lists and rows contained in list and rows and thus making it possible to hold JSON compound objects and arrays, or lists/rows hierarchies for example.

The code of this project is aimed to generate compiled COM (ActiveX) dlls so the three main classes CList, CRow and CMapStringToLong can be reused at maximum speed in other projects and/or COM compatible languages.
Anyway, instead of using an external DLL in VBA/twinBasic projects, just adding the following core classes source code files, to a VBA or twinBasic project, gives you the benefits of these classes without having to register a COM dll in your system:
 - CList.cls, CRow.cls, CMapStringToLong.cls, IListCompare.cls, MRowList.bas

There are no other dependencies, except the ones of your languages (VB5/VB6 runtimes), or none when compiling with twinBasic.

All the other files are here to support testing.

I still use Microsoft Access as my main VB/A development environment, so you'll find here two accdbs that I use to maintain the project:
- RowListLibFull_Tests.accdb is the one containing the whole enchillada;
- RowListLib_Tests.accdb is (or should be) the same, but it uses the twinBasic compiled 32bits binary of the library instead of the core classes source code, so tests can be run one the compiled library.

The code has been made compatible with twinBasic recently, and a 32bits binary version, that I actually use in production applications, can be downloaded from my company's website: [RowListLib_win32.dll on devinfo.net](https://devinfo.net/download/RowListLib_win32.dll) .

# Oops

Check if there's a "Stop" statement in the CompNonObjectValues() method of CList and please delete the line or comment it; although I never reached the code in my tests and in production, that is an application killer if that would happen. I'll get it fixed but in the mean time, sorry for that, you'll have to eventually do it yourself.

## A bit of history

I started to develop this library a long time ago. My inspiration dates back to 2003, when I came to experience the list object available in the OMNIS L4G cross-platform development environment. Since then, I never came across anything like it, so I continued developing these classes and using them to boost my productivity with quite some success in RAD and production uses cases. I cannot count how many of my classes and solutions are based on these classes, so hopefully they're not too buggy, at least they're stable enough for me to use in production.

### Porting C# code to VB/A

I have to say that without these classes, I would probably never succeeded porting [codebude's QRCode .NET library](https://github.com/codebude/QRCoder).
I needed a library to generate Swiss QRcodes, and back in 2020 I decided to "port" this one in VB5.
In short, and with all due respect for the work and the author who I'm grateful for putting out one the nicest and freely resuable solution out there, this was nonetheless an unbelievable *nightmare*. Worst than anything was transforming LINQ expressions reliably, which anyways was made easier by extending the CList class with the ListGroupBy() methods. Performance was an issue in developement environments (Access) where generating a Swiss QRCode takes around 30 seconds, while compiled (actually as an ActiveX EXE server) takes around 1.5 seconds, without much optimisation.

If not already done so, take a look at my repositories to find the port soon.

## Technical notes

The list and row library contains two main VB/twinBasic classes, CList and CRow which are general purpose, multi-column list (CList) and row (CRow) classes.

CList is in fact a (big) wrapper around a two dimensional array, contained in a Variant.
Basically, it allows the class user to see the internal array as rows and named columns, just like a (DAO/ADO) recordset does, although recordsets are not entirely held in memory (but the CList and CRow internal arrays are).

To match a column name with a column index, a CMapStringToLong class instance is used internally by each object.
CMapStringToLong is like the part of a VB/TB Collection that matches item names to objects, but instead on using a hashing algorithm, it uses dichotomic search in a string array.

While maintaining column names and definitions in list and row objects is certainly a significant (but not excessive) overhead compared to an array, it brings an astonishing and very appreciable flexibility to a simple two dimensional array, far from the so basic and slow Collection and/or Dictionary classes, which by the way, god only knows where to find any of the source code.

Take a look at the test code to see how many ways there are to define a list or a row and to add/delete/sort rows.

### Performance

There's some interesting tests on performance where insertion and retrieval is compared to the performance of a collection.
CList being an array, better performance is achieved by growing the internal array by more than one element when needed.
Deleting rows does not remove them from memory, instead, there's a parallel array storing the indices of the row vs their real internal array index. While this adds indirection to access the array elements, it pays on deletion and sorting, because only the parallel array is access directly and the Win32 CopyMemory (actually Rtl_MoveMemory, as you probably know) can be used to move contiguous multiple elements directly.

### Database

Wrapping ADO in a standard module API and integrating CList and CRow instead of keeping opened recordsets has been a time saver and productivity booster for me since a while. You can find a version of the API I use in MADOAPI.bas.
The function I use the most out of this API is ADOGetSnapshotList().
The trick is to use an ADO recordset temporarily and get all the rows in one operation with ADO's GetRow recordset method which conveniently places the whole recordset in a two dimensional array, contained in a Variant. That is similar to the internal array leveraged by the CList class, so all that is remaining to do is to give that to a CList instance, possibly without copying it. I did not find a better way than exposing the CList internal data array for that, which totally breaks encapsulation, but solves the problem.

### Gotcha(s)

There's an heavy performance hit, if you create and define CList (respectively CRow) objects inside a loop, as the time to build the internal structures holding the column names and types then takes a lot of time.
To circumvent this problem, you can create and define a CList instance outside of the loop and use the reset method inside of the loop to empty it without destroying the columns definitions.
I recently created updated MADOAPI and created a CList memory pool manager to address this problem further and will post it either here or [on my blog](https://francescofoti.com) where you can follow me (you'll also find me [@unraveledbytes](https://twitter.com/unraveledbytes) on twitter).


