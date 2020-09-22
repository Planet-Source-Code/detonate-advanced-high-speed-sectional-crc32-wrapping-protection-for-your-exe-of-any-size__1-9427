Sectional CRC32 Protection (High-Speed) for your Exe
by Detonate (detonate@start.com.au), for PlanetSourceCode.com

You are free to use this in your commercial/uncommercial/whatever
programs, and I dont for any credit, but if you do publish this
source code, please keep it intact and unmodified, that is all i ask.

The idea for this stemmed because the original CRC32 wrapper
i released at planetsourcecode.com/vb which used the entire file
was too slow for practical use if the exe was over a few hundred
kb in size. Its a relatively solid protection, as its based on
every byte in the file, meaning its virtually impossible to change
any bytes without changing the hash.

Speed was the problem. If you dont understand why, try this:
    For I = 1 to 5000000
     Doevents
    Next I
Thats looping through just 5 megabytes, and remember its doing
NO processing either!  yet its still slow... so if we cant even do
that, how can we get an effective CRC32 wrapper over all the bytes?
... buggered if i know! :-)

So i put my mind down a bit... and thought... we only really need to
protect certain parts of the file... and if those parts of the file
included the Exe header, the CRC32 instructions, the checksum , and
various other fields, then the protection would be just about as good
as one applied on the full file. The only catch is that you have to
identify which parts of your exe are important... this is dead easy
if you have a hex editor.

Enter Sectional CRC32 Protection!
This version only uses the sections of your .exe that you deem important
For example, there is probably no need to protect a BMP image in your
exe file, but you may want to protect where it says "UNLICENSED" :-)
To do this, use a hex editor to find the byte entry point, and then
figure out how many bytes you want to protect... whether its 1 or 100000
or whatever. You add this string to the CRCTest string, which holds
all of the "important" strings.
After youve filled CRCTest with all of your important strings, you
create a CRC32 hash on THAT, and then append that to the end of the
file.

This example uses two "protected fields" - the first 500 bytes of the
file, ie. the exe header, and also a second imaginary field which could
contain whatever, just to demonstrate how to select fields. Because
its reading directly from the disk and not from memory, it doesnt matter
if youre file is 1 meg or 100 megs.. the file contents are never read
into memory
