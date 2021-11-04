# MYANMARCalendarXL
Routines for Calculation and Conversion between other calendar(s) and Myanmar Calendar as UDFs...\
As of 04NOV2021, the following usages are possible.\
Please note that the outputs of the functions are of type: <b>TEXT in Excel/String in VBA</b>.
![as UDF](/images/MMRCalXL_in_Excel_asUDF.png)

## Brief background
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

## Bloopers
This part is reserved for the kinks that I faced while porting in from the Javascript code.

I found that VBA's own mod function providing the modulus of a division, returns Integer types as result and that caused an error in calculation which took me 3-6hrs or so, to finally found it as the culprit (also because I don't fully understand the idea behind the whole Myanmar calendar concept and my unfamiliarity with Javascript).\
It doesn't seem very big but when <b>3 mod 2.5</b> produces 1 where as a floating point modulus function would return <b>0.5</b> which would causes a chain reaction of adding up days like 2weeks to 50days in a calendar calculation!

The second bottle neck I encountered during the porting process was setting a variable that I supposed should hold a boolean which I didn't know or forgot how Javascript handles booleans thus causing it to hold a VBA value of -1 equivalent to the Boolean TRUE while the original Javascript code assumed it to hold just a 0 and I was off by something like a staggering 4-5 months into the future.\
Phew! Thank god I found it after 6hrs of being stuck trying to solve it by pulling my hairs out (now I just shaved my head like an egg. I mean, why not?!).

The last issue is that there no intrinsic mechanism implemented in the original functions to check/reject the user from entering a Oo/Hnaung value (in English, Early or Late) Tagu(?Kason) month information while lacking it in parameter passing can cause errors in Burmese Dates into Western Dates conversion.\
From Mr. Yan Naing Aye's page, I thought that it only occurs in the Tagu or Kason months and I believe that Tagu has more chance. It was clearly explained on his page and many thanks for sharing that knowledge.\
But I believe that if we don't provide it, it might cause errorneous calculations. Need to check this further.\
The best method I found so far is: to find the Burmese Date of a western date in question first, then using the output of that function=the correctly configured string for passing into the Burmese->Western Date conversion function because if a misconfigured Oo/Hnaung information were passed (the possible values are just a 0 and 1), there will be some wrong returns from the B->W conversion function.

### Requesting permission to port Javascript code to VBA code
I have submitted a request to the original author: Mr. Yan Naing Aye to allow me to convert his Javascript code into VBA from his website's comment feature (that was like a week ago) and today through his LinkedIn page and so far (as of 04NOV2021), no reply was received yet.\
Until I am officially allowed to port, I won't be able to share my VBA code.

## Acknowledgements and Thanks
All credits goes to Mr. Yan Naing Aye and also to the people who asked a simple question like, "How do we convert Myanmar dates to Western Dates...and vice versa?"...|
We porter are literally like the Porters who carry stuff from the sellers, in this case the original author and the buyers, the Myanmar Users, to their doorsteps...

## Further Information
I searched and found that Mr. Yan Naing Aye has also a GitHub repo [here](https://github.com/yan9a/mmcal).\
I found that there are both Javascript version and C++ version over there but the Javascript code is more complicated for me there, so, I just stick to the more understandable code and explanations on his website.\
And there's an awesome realtime interactive calendar developed and shared by the same benefactor over [here](https://yan9a.github.io/mmcal/).

## Things to do
- Improve the current VBA code into a more streamlined, error-free and clean and elegant code
- Will probably allow the parameters in and out of the UDFs in Burmese language/font and if this happens, it will be limited to Unicode fonts only (Myanmar people, please stop using Zawgyi font).
- Will probably take the pains to come up with an Excel formula but afraid that it will be quite staggering considering the calculations involved nonetheless, that's the first thing on my mind as of now
