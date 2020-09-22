_________________________________________________________________________________________________

»Ripper Program« - By Alex M
_________________________________________________________________________________________________


_________________________________________________________________________________________________

Notes
_________________________________________________________________________________________________
	It saves Extracted/Ripped files to "C:\" as a Default, It Also Saves Extracted Viv files to "C:\RippViv"

_________________________________________________________________________________________________

Comments
_________________________________________________________________________________________________
	I am a bit dissapointed on one comment I received on this program, it said "Good program, but it seems that posting your code is more important to you than documentation". I have had tons of work for uni, with exams comming up next week, I have been studying and have had no time. The reason I posted this code at this point in time was because I wanted to show people what you can do with Vb, I do not care about ratings, All I care is that people benefit from my program. Anyway, here is the documentation despite my lack of time due to my exams. All I have to say To the person who made the comment, Thanks a lot for putting my 6 month project down in your 15 second comment!!
	Anyways, To the people who thanked me for a well done program, Thank you for showing your support and kindness, I will try and post more of my past projects up after my exams or when I have some time.

_________________________________________________________________________________________________

Description
_________________________________________________________________________________________________
	This program can search a file for other files that it contains. It searches for Bytes that are common to the file format, such as .Bmp files have "BM" as their first two characters, and have their file size in their header. While sometimes there are characters that match these in a file, most of the time, you will be able to ripp files correctly (NOTE the "Delete corrupt or incomplete Bmp and Gif files" is not always accurate, you sould use a program like Paint Shop Pro to determine whether they are Bmp or Gif files).
	The Expand Viv option allows you to extract files from an Electronic Arts .Viv Archive. The files it contains are usually .Qfs or .Fsh, these files can be opened by a program like Nfs ToolBox (Although it Doen't support NT), or Ifranview (A really good picture viewer for almost every format).
	The Search For Text in Exe option, allows you to get all text data from within a file. This can be useful if your friend has given (or not given) you a vb program with a password and you know it is stored in the program. To those users who are now fretting about how easy passwords can be found in Vb projects, you can easily 'mix' up the letters in a string, then make a subroutine to un-'mix' them (Do not Store non-encrypted Passwords in you VB project, i.e if text1.text="password" then... , It can be found easily and with an accuracy of 100% each time!!)
	The encryption option allows you to encrypt files using a simple matrix encryption technique. The concept is simple, it uses the password as a seed, Therefore the bigger your password the better (just don't forget your password!!). Ofcause if you know the first say the first 10 characters of an encrypted file and you know how it is encrypted, then you can work out the first 10 characters of the password (But most users would not know how to anyway, ie your brother, sister or friends). Also, I will not make any sort of password hacking tool, It would be stupid to code an encryption technique then make another program to easily undo it (But it can be done)!!
	The replace file option allows you to replace a File within another file, say a splash screen for a program. I, myself have patched the Adobe splash screen, All you need is the same format and the size or smaller picture, then replace (But you need to know the position of the picture within the file i.e. it is the 6th picture or the 7th picture e.t.c.)
	You can use any part of my code that you want, I hope you get the most out of using my code and/or program. Have Fun!!

_________________________________________________________________________________________________

Difficulties
_________________________________________________________________________________________________
	"All the files I extract are corrupt!!" - Thats probably because they aren't actually files, they probably are just data that the program though it was a File.
	"Every time I replace a file, The original becomes corrupt!!" - Thats because most programs know exacally how big certain files are ment to be and the exact file format, and when you replace them, your picture may be of a different size or of a different format, also check the colour depth (i.e. there are several different compression techniques for .Jpg files, as there are for many files).
	"I want to add file To a .Viv file as well as extract!!" - I suggest downloading 'Viv wizard', but I have made another program to Extract and Insert files into and out of a .Viv file (It will be posted much later this year).
	"My computer does not shutdown properly when I select 'shutdown when done'" - I know, It seems to only not work on certain computers, You will need to find some code on the ShutdownWindowsEx Api call (I suggest finding a copy of AllApi Guide on the net).
	"I can't find a correct passord in a Exe file" - Thats because the password is not always stored in the exe file itself, or it is encrypted, or it could even be generated (i.e. a Serial key). In these cases, you will have to look somewhere else for a crack or a serial key database (There are plenty on the net, but I do not endorse them or what they do, Thats how most people get viruses by running dud cracks from the net, I do not like the people who make viruses just to annoy others)
	Any other query's, just E-mail us at Alex_murray1@hotmail.com (make sure that the subject doesn't contain '..' or 'adv', also make the subject appropriate or I'll think it is Spam)
	



