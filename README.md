# 1.MyanCalXL
Routines for Calculation and Conversion between other calendar(s) and Myanmar Calendar as UDFs...\
As of 04NOV2021, the following usages are possible.\
Please note that the outputs of the functions are of type: <b>TEXT in Excel/String in VBA</b>.
![as UDF](/images/MMRCalXL_in_Excel_asUDF.png)

# 2.Brief background
I am a member of several Facebook groups where people discuss and ask for help on topics related MS Excel.\
In those places, most of the frequent questions can quite easily be solved by Excel formulas, but many a question popped-up from time to time that managed to wake the curious cat in me to wonder whether to use VBA to solve them.\
One of the many of these questions is: "How do I change from Myanmar date to English Date?" or vice versa...\
The usual answer is "Add 638 to Myanmar year!" which is the only method for Myanmar Calendar conversions that I have known in all my years.

So the desire to help them as well as to quench my own selfish thirst for knowledge led me to search for solutions to this particular question.

Then one day, I found this [webpage](http://coolemerald.blogspot.com/2013/06/algorithm-program-and-calculation-of.html?m=1) by Mr. Yan Naing Aye.\
On the first read through, I was amazed and humbled by the awesomeness alone.\
The topic was very complicated for me at that time but the person behind it was extremely knowledgeable on the subject matter and explained everything quite clearly as I found out on later readings.\
I have other projects going on at the time I first found it, so I had to drop it for some time.\
Then when I can finally work on it, I have to read through everything again. Even now that I finished porting the Javascript on that page into a fully functional (yet pretty raw) VBA code, I am not ashamed to admit that I still do not fully understand the theory behind it.

I write code in VBA for Excel just because I love VBA but not because I was trained to write VBA code (or any) and my Javascript knowledge is only part-way through a codecademy free course. Therefore, trying to understand the data structures and the code mentioned is not very easy for me, however, I managed to finish it in about 1 week's time.\
This is not an excuse nor gloating what I have done but just recording what I have encountered. I have to say it out loud that, it was not hard but it just was not a walk in the park either.

I have just simply ported Mr. Yan Naing Aye's Javascript code to a working VBA code to be used in MS Excel. The original concept, theory and algorithm behind this whole thing belongs to Mr. Yan Naing Aye.
I ported it all to VBA just to help our people (oppressed and terrorized by a group of military generals who are cruel, senseless and immoral beyond anything - still true, as of 04NOV2021) who are very resilient, content but thirsty for knowledge yet taken-advantage-of, misled and scammed (on facebook and off) by countless so-called IT tutors who claimed that they can teach all about excel in a week with fees sometimes as low as USD5 (imagine Stilgar's spit scene in the meeting with Leto Atreides, god forbid!)... 

This was taken yesterday on 03NOV2021.
![just finished](/images/MMRCalXL_in_ImmediateWindow.png).

The current code is still pretty raw in that it was barely functional after being almost directly ported in from their Javascript counterparts.\
I will have to develop it further to make the code more efficient and the code leaner and more like VBA.\
And when it is ready, I will release a package here.

## 2.1.Bloopers
This part is reserved for the kinks that I faced while porting in from the Javascript code.
### 2.1.1
I found that VBA's own mod function providing the modulus of a division, returns Integer types as result and that caused an error in calculation which took me 3-6hrs or so, to finally found it as the culprit (also because I don't fully understand the idea behind the whole Myanmar calendar concept and my unfamiliarity with Javascript).\
It doesn't seem very big but when <b>3 mod 2.5</b> produces 1 where as a floating point modulus function would return <b>0.5</b> which would causes a chain reaction of adding up days like 2weeks to 50days in a calendar calculation!
### 2.1.2
The second bottle neck I encountered during the porting process was setting a variable that I supposed should hold a boolean which I didn't know or forgot how Javascript handles booleans thus causing it to hold a VBA value of -1 equivalent to the Boolean TRUE while the original Javascript code assumed it to hold just a 0 and I was off by something like a staggering 4-5 months into the future.\
Phew! Thank god I found it after 6hrs of being stuck trying to solve it by pulling my hairs out (now I just shaved my head like an egg. I mean, why not?!).
### 2.1.3
The last issue is that there no intrinsic mechanism implemented in the original functions to check/reject the user from entering a Oo/Hnaung value (in English, Early or Late) Tagu(?Kason) month information while lacking it in parameter passing can cause errors in Burmese Dates into Western Dates conversion.\
From Mr. Yan Naing Aye's page, I thought that it only occurs in the Tagu or Kason months and I believe that Tagu has more chance. It was clearly explained on his page and many thanks for sharing that knowledge.\
But I believe that if we don't provide it, it might cause errorneous calculations. Need to check this further.\
The best method I found so far is: to find the Burmese Date of a western date in question first, then using the output of that function=the correctly configured string for passing into the Burmese->Western Date conversion function because if a misconfigured Oo/Hnaung information were passed (the possible values are just a 0 and 1), there will be some wrong returns from the B->W conversion function.
### 2.1.4
After the launch of this repo, a friend and a colleagure, Mr. Sithu Kyaw, shared his opinion that it would be nice to have the output (and/or input) as Burmese Unicode font.\
And for 2 days, I worked on it. And yesterday, it was achieved (Well, the output part!).

It was something I thought of even during the original porting process near the end. However, the only limitation that stopped me from going in that direction during the initial porting process was that, the VBE/VBIDE or simply, the VBEditor window is not compatible with Unicode font as it is made for ANSI code pages or so they said. And if I create constants like MMRMonths="တန်ခူး,ကဆုန်,..." etc., the VBE will only show ??????s and worse still, output to worksheet as ?????s.

Of course, we all know that we can output unicode strings to the worksheet and work with unicode strings in memory but the limitation only applies to the VBE.\
So, my first attempt to overcome that problem was finding ways to be able to declare unicode constants in VBE by changing the Non-Unicode friendly programs to use Burmese/Myanmar from the Control Panel's Region/Locale settings or use some apps like LocaleEmulator/NTLea.

The problem with this approach is that if I do that, I can write unicode strings in VBE but not every user would be willing to install new apps or restart their computers after changing the said Locale settings.\
Thus, if they open the .xlsm's project in VBE, all the unicode strings I declared will turn into ?????s like all those people in Midnight Mass come morning and irretrievably gone forever.

Therefore, I decided to convert the Myanmar month names into their own respective character codes using ascW funtion in VBA so that "တန်ခူး" becomes "4112|4116|4154|4097|4144|4152" and hardcode it into the code to be distributed to the user and convert it back to "တန်ခူး" again at the start of the function.\
It's messy I admit but hey, it works! Here, I learnt something that if I use split and the join to create the unicode string, it won't put back the unicode string together and I must use & to rejoin the converted unicode parts. It's strange to find out the & and the join don't function the same!

Another issue that came with outputing Burmese fonts is the numbers. The numbers are not encoded in Unicode so they can just be converted to Burmese numbers once they are already in Worksheet by setting the UDF's parent cell's font as "Pyidaungsu Numbers". Without doing that step, if you choose to output to Burmese font, you will get a mixed English and Burmese output not unlike "1383/နတ်တော်/လဆန်း/4". Since UDF's cannot change the Formatting of the Excel UI, I cannot help with that, sorry.

Enough said about Burmese outputs but one last thing remains that I, for one, do NOT wish to output anything except Numbers from the UDFs. For me, the numbers are the truest, cleanest and most efficient outputs from a UDF, that can be Matched, LookedUP, used-with-Data-Validation-dropdowns at the whims of the User. But in the interest of the ease of use of/by my beloved Myanmar people, I made this effort.\
Epilogue: one day, I might work on Burmese inputs to the UDF: toMMRDate...\
![Burmese output combined](/images/MyanCalXL_Burmese_Unicode_Combined.png)

# 3.Requesting permission to port Javascript code to VBA code
I have submitted a request to the original author: Mr. Yan Naing Aye to allow me to convert his Javascript code into VBA from his website's comment feature (that was like a week ago) and today through his LinkedIn page and so far (as of 04NOV2021), no reply was received yet.\
Until I am officially allowed to port, I won't be able to share my VBA code.\
<b>As of 04NOV2021-1800, Mr. Yan Naing Aye has kindly and graciously allowed me to port his Javascript code on his [webpage](http://coolemerald.blogspot.com/2013/06/algorithm-program-and-calculation-of.html?m=1) into VBA code.</b>\
However, I don't feel like the current VBA code is ready as of now to be distributed to the public as I need some time to clean up the code and review it to make it better and operate more efficiently.

# 4.Acknowledgements and Thanks
All credits goes to Mr. Yan Naing Aye and also to the people who asked a simple question like, "How do we convert Myanmar dates to Western Dates...and vice versa?"...\
We porters are literally like the Porters who carry stuff from the sellers, in this case the original author and the buyers, the Myanmar Users, to their doorsteps...

# 5.Further Information
I searched and found that Mr. Yan Naing Aye has also a GitHub repo [here](https://github.com/yan9a/mmcal).\
I found that there are both Javascript version and C++ version over there but the Javascript code is more complicated for me there, so, I just stick to the more understandable code and explanations on his website.\
And there's an awesome realtime interactive calendar developed and shared by the same benefactor over [here](https://yan9a.github.io/mmcal/).

It was my fault to add this information late because at the time of writing this README, I forgot to add a source that I have read to understand how Myanmar Calendar could be calculated in addition to Ko Yang Naing Aye's blog. Even while it has not been directly used in the porting of JS code to VBA but more like important to understand how Myanmar Calendar was calculated. This source is [မြန်မာပြက္ခဒိန်တွက်နည်း](http://shwenyein.blogspot.com/2012/07/blog-post_1862.html) by U Aye Nyein. It seems like Ko Yang Naing Aye himself consulted U Aye Nyein in the development of his own algorithm. In any case, I would like to thank U Aye Nyein for sharing his knowledge which is a very valuable source on the history of Myanmar Calendar.

# 6.Things to do
- Improve the current VBA code into a more streamlined, error-free and clean and elegant code
- Will probably allow the parameters in and ~~out of the~~ UDFs in Burmese language/font and if this happens, it will be limited to Unicode fonts only (Myanmar people, please stop using Zawgyi font).
  - Output to Burmese Unicode Font viz. Pyidaungsu Numbers was done as of 06NOV2021
  - ![Burmese output single](/images/MyanCalXL_Burmese_Unicode2.png)
- Will probably take the pains to come up with an Excel formula (non VBA-based-UDF) but afraid that it will be quite staggering considering the calculations involved nonetheless, that's the first thing on my mind as of now

# 7.Wiki Pages
Since the parameters/arguments that could be passed to the various functions, I believe that it is better to create a wiki on that matter so that the end-users can refer to the wiki pages rather than having to go through the extra-long Readme file here. Therefore, for more information on how to use the date conversion functions, please refer to the [wiki pages](https://github.com/4R3B3LatH34R7/MYANMARCalendarXL/wiki).

# 8.Releases
I believe that a brief explanation is required why the MyanCalXL is going to be released closed source.\
Recently some people took advantage of my MMRTokenizer code and they used the ideas and concepts I openly shared with public to their advantage without recognizing the efforts I put into that project.\
I don't really care about being recogized/credited but I didn't work hard for that person's benefit.\
Therefore, all the future releases of my projects will be closed source.\
Another reason is because of the fact that the complexity of the arguments passed to the functions demand help documentation on those arguments.\
The UDFs are best used with a contextual helps to make the most out of them.\
Therefore, I am going to release a .xlsm version which has Excel-DNA Intellisense embedded and the users only need to install Excel-DNA-Intellisense addin.

Kindly check [Releases](https://github.com/4R3B3LatH34R7/MyanCalXL/releases)
NB: Please download .chm together and install [Excel-DNA Intellisense .xll](https://github.com/Excel-DNA/IntelliSense).

## 8.1.Initial Release v1.0a
1. [v1.0a](https://github.com/4R3B3LatH34R7/MyanCalXL/releases/tag/v1.0a) released on 01JAN2022 1100 MYANMAR STANDARD TIME.\
<b><ins>This release is dedicated to the Brave but Gentle Myanmar people who are fighting back the brutal fascist military regime.\
  Viva la revolución!</ins></b>

# 9.Format
The format of describing the Myanmar calendar date was set to follow the Myanmar newspapers & TV news announcements which I supposed were setting the basis for official writing/declaring/describing format of Myanmar date and time.\
Reference:At the end of the [WikiWand article](https://www.wikiwand.com/en/Burmese_calendar#/overview), under the section, "Official formats".

***
## License
I don't actually like/want/wish to apply CC BY-SA license to what I share, really!\
However, there exists some jerks in this world who thought it's ok to derive my work without proper accreditation.\
I don't care much for fame nor finance but a little credit for the many hours of my limited life I spent on a project is appreciated.\
Shield: [![CC BY-SA 4.0][cc-by-sa-shield]][cc-by-sa]

This work is licensed under a
[Creative Commons Attribution-ShareAlike 4.0 International License][cc-by-sa].

[![CC BY-SA 4.0][cc-by-sa-image]][cc-by-sa]

[cc-by-sa]: http://creativecommons.org/licenses/by-sa/4.0/
[cc-by-sa-image]: https://licensebuttons.net/l/by-sa/4.0/88x31.png
[cc-by-sa-shield]: https://img.shields.io/badge/License-CC%20BY--SA%204.0-lightgrey.svg
***
 <a href="https://trackgit.com">
<img src="https://us-central1-trackgit-analytics.cloudfunctions.net/token/ping/kybbjinclq8j02pyd4ik" alt="trackgit-views" />
</a>
