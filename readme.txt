BldRnd.EXE v1.3

The slowest part of the process is the initial creatation of the 32kb
random data string.  This is built for every file, regardless of size.

 23-DEC-1999  1.0  Module created by kenaso@home.com
 30 SEP 2000  1.1  Added the available drive letters
 10-DEC-2000  1.2  Added some common routines and added some
                   more documentation
 02-APR-2001  1.3  Removed two classes.  Added a class for access
                   to Scripting.FileSystemObject (scrrun.dll) and 
                   updated random data generation class.  Added 
                   shutdown switches that stop processing almost 
                   immediately.  Created single data string
                   of 32768 bytes of random data in which to build
                   test files.  1 mb test file is now created in under
                   15 seconds on a 500 mhz machine.                   

Build random generated test files.  These test files would 
normally be used to test data transfer rates, fillers for
disk capacities, etc.  All files are created in the drive C:
root directory.

The main purpose of the program is to build a test file to your
specifications.  

o	The file can have fixed length records or one long record to 
    match the size of the file.

    The fixed length record will be 78 printable Ascii Text 
    characters (decimal values 33 to 126) and the trailing two 
    characters for each record is the carriage return (chr(13))
    and the the linefeed (chr(10).  Without a hex viewer, these
    will appear as blanks.
    
o	You can choose between predefined lengths or you can design 
    your own file.  The predefined lengths are listed in the combo
    box.  If you opt to customize a file, the combo box changes to
    a user input box. 
    
o	You have the option to use all ASCII characters (0 to 255) or
    just the keyboard printable characters (33 to 126).
    
o	You have the option to leave these characters as a single entity
    or convert them to their hexidecimal two character representation.    

o   Use memory cache to build long data strings.  Much faster.

-----------------------------------------------------------------
Written by Kenneth Ives                    kenaso@home.com

All of my routines have been compiled with VB6 Service Pack 4.
There are several locations on the web to obtain these
modules.

Whenever I use someone else's code, I will give them credit.  
This is my way of saying thank you for your efforts.  I would
appreciate the same consideration.

Read all of the documentation within this program.  It is very
informative.  Also, if you learn to document properly now, you
will not be scratching your head next year trying to figure out
exactly what you were programming today.  Been there, done that.

This software is FREEWARE. You may use it as you see fit for 
your own projects but you may not re-sell the original or the 
source code.  If you redistribute it you must include this 
disclaimer and all original copyright notices. 

No warranty, expressed or implied, is given as to the use of this
program. Use at your own risk.

If you have any suggestions or questions, I'd be happy to
hear from you.
-----------------------------------------------------------------
