<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->
<%
if not LoggedEmployee then Redirect("login.asp?Source=tips.asp")
%>

<p align="center"><span class=Heading>Selling Tips</span><br>
<span class=LinkText><a href="login.asp">Back To Salesman Options</a></span></p>

<p align="center" class=SubHeading>
Things to remember.
</p>
<p>
1.  Above all, there is one thing you must remember when selling GroupLoop - <b>they need this</b>.  This is a tool that 
can completely revolutionize the way churches communicate.<br>
2.  For the most part, churches can easily afford this.  Stress that it can cost as little as $40/month with no setup fee.<br>
3.  Sales can become very frustrating.  Do not get discouraged quickly.  It can take some time before you find your sales angle.<br>
4.  Be as kind as possible - even if others aren't.<br>
5.  <i>Know what you are selling.</i>  Below will tell you how to get aquainted with GroupLoop.
6.  PRINT OUT OUR FLIER.  Hand it out to everyone.  You can get it <a href="material.asp">here.</a>
</p>

<p align="center" class=SubHeading>
Learning more about GroupLoop
</p>
<p>
Handing out a flyer is only part of the sales process.  People will usually ask you questions.  The best way to become familiar with 
our sites is to actually use one.  Goto <a href="http://www.GroupLoop.com">GroupLoop.com</a>.  Read all the available information, 
and pay attention to all the features our sites offer.  Then go ahead and sign up for a free site.  Fool around, and you will quickly 
become accustomed and will be able to answer everyone's questions.  If the person has a question you can't answer, tell them 
to e-mail support@grouploop.com.  Or ask us yourself.
</p>

<p align="center" class=SubHeading>
A Sample Sales Pitch
</p>

<p>We have written a small, standard sales pitch that you could start with.  Do not try memorizing this word for word.  It will 
just give you a better idea.
</p>

<p>
"Hello, my name is <%=GetEmployeeFirstName(0)%> and I am from GroupLoop.com.  GroupLoop is a company that specializes in making 
web sites for churches like this.
</p>

<p>
Until now, the only sites churches had were just informational - here's our history, location, 
and phone number.  GroupLoop took it way beyond that.  They created a system that will make communication within your church much 
more efficient.
</p>

<p>
As you probably know, communication in a church is tough.  All the different committees, made up of different people.  
Members of the church have different lives, and can't always be here.  GroupLoop created a system that allows each committee to have 
their own web site.  Members of each site can add announcements, meeting minutes, discuss things, vote, post photos, files, and lots 
more.  This is all simple to do, and noone needs to know anything about web design.
</p>

<p>
Your main church site will be able to provide 
visitors and members with the latest information - church times, announcements, even post of up your sermons.
</p>

<p>
And it's REALLY cheap.  
There is no setup fee, and the monthly cost will run you anywhere from $40-$100, depending on the number of sites you get.  Some 
churches have over 20 different sites.  This really will change your church.  Here, take this flier to learn more about GroupLoop."
</p>


<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
