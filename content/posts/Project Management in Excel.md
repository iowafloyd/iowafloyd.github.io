+++
date = '2025-09-11T18:08:43-05:00'
draft = false
title = 'Project Management in Excel'
categories = ['coding']
tags = ['project management', 'excel', 'gantt']
+++

## Why project management?
The curriculum at Iowa State University taught me a lot (statics, mechanics, thermodynamics, procrastination...), but it didn't teach me anything about project management. Why is that? I've had to use some form of project management in my daily professional life since day one.

I'm betting that project management is just one of those things that's hard to put into a textbook. Because (at least in my experience) you only really learn project management when Plan A goes sideways. So you make Plan B... half of which goes sideways. So you make... You get it.

I've worked for three companies thus far in my career, but by and large most of that tenure has been spent with just one company, Viking Pump. Most people haven't heard of them, though they've been in business since 1911, their products pump things you use everyday but never think about (toothpaste, asphalt, marshmallow fluff), and they're one of very vew companies that can truthfully say they have products on all sevent continents. Let that sink in for a moment.

As I've progressed in my career at Viking, I've taken on more responsibility with the overall ownership of a project. At first, I was just responsible for the CAD modeling and some calculations. Further on, I took on the development testing and validation stage. Further on, I took on the total pump design, including the entire bill of materials, working with Supply Chain to get new parts quoted and ordered, and owning the total cost of the product. Finally, added onto this was the ownership of the whole process: everything above, but also planning out just who has to do what, when things need to be done, and when the concept has been fully tested and proved out, and is ready for production implementation.

That's a lot to keep track of! And all this is done for each project, where each engineer handles a number of projects simultaneously.

## Brain power: low-gear vs. high-gear
All these... "logistics"... are difficult for my brain to handle. My brain just isn't as efficient at it as it is with details. Or in more mechanical car-related terms, my brain is most efficient when it's in gears 1 through 3. The low-speed, high-torque gears. Chugging away solving a technical problem with lots of little variables all at play, but somehow there's a path forward if you just get some traction bit-by-bit.

Project management, on the other hand, is more like gears 3 through 5 (or 6, if you've got one of thoes fancy cars). I see project management as the ability to get from here to there -- that far-off, distant goal -- and mapping out the quickest, most efficient route. You'll need traction along the way, but gears 3 through 5-or-6 are all about speed and efficiency.

The difference between cars and people in this analogy is that people tend to be better at one or the other. My brain-engine is very efficient going from fully-stopped to rolling along at 40MPH in 3rd gear. My brain-engine has a hard time transitioning into 4th and 5th gears to really get speed. If my brain is kept in these high gears for too long, like any mechanical device stuck in a low-efficiency state, it starts to overheat and not work as well.

So, I've had to lean on tools to help me do the job. The tool I needed was one that allowed me to keep track of both the little detailed things, and the larger logistic things, all in one place. And for me to be able to "see" things further out in the future, it had to be fairly visual. It seemed what I needed was a good ol' Gantt chart.

## Inspiration: Excel Gantt v 1.0
I didn't make this from scratch; I actually found a much simpler version online when I'd searched for "Excel Gantt chart" or something similar. The original version was created by [Glenna Shaw](https://www.linkedin.com/in/glennashaw/), but I can't find the original version to display here.

Glenna's version could track one project, with a place to input tasks and start/end dates, who the owner of that task was, and even information on how much they were paid by the hour, I think. Then, as outputs, there was the Gantt chat and also a total project cost ("are we ahead or behind" kind of thing). It was pretty well laid-out, I have to say. It was a great place to start, but I needed more multi-purpose functionality. I didn't want to have separate spreadsheets for each project. I wanted a one-stop-shop for everything that I was working on, whether it be a years-long platform project or a couple-days administrative task. Only when I could track everything from the big to the small, and have it all laid out in a timeline, could I use this tool to get better at **thinking** like a project manager. That was the key.

## Excel Gantt 2.0: the "Project Master"

Below are a few screenshots showing the major functionality of the spreadsheet. I put a lot of "UX thinking" into the layout, where the main page is purely for visual data output and all other pages are for data input. An Excel version of "front-end vs. back-end".

I'll explain more about each screenshot in next few days, but for now I just wanted to get this blog post mostly written and up on the site.

### Project Timeline
<a href="/img/01_Project_1_Timeline.png">
    <img src="/img/01_Project_1_Timeline.png" alt="Screenshot of Timeline view - the Gantt chart" />
</a>

(more info about this screenshot to come later...)

### Action Items
<a href="/img/02_Action_Items.png">
    <img src="/img/02_Action_Items.png" alt="Screenshot of list of action items" />
</a>

(more info about this screenshot to come later...)

### Project Drop-down
<a href="/img/03_Project_Dropdown.png">
    <img src="/img/03_Project_Dropdown.png" alt="Screenshot of drop-down list for changing projects" />
</a>

(more info about this screenshot to come later...)

### Miscellaneous Timeline
<a href="/img/04_Misc_Timeline.png">
    <img src="/img/04_Misc_Timeline.png" alt="Screenshot of Timeline view of the Miscellaneous items" />
</a>

(more info about this screenshot to come later...)

### Inputting Action Items
<a href="/img/05_Inputting_Action_Items.png">
    <img src="/img/05_Inputting_Action_Items.png" alt="Screenshot of inputting new action items" />
</a>

(more info about this screenshot to come later...)

### Calendar View & Date-Picker
<a href="/img/06_Calendar.png">
    <img src="/img/06_Calendar.png" alt="Screenshot of yearly calendars from 2025 to 2030" />
</a>

(more info about this screenshot to come later...)

### Auto-Rebuilding of Timeline
<a href="/img/07_Misc_Auto-Rebuild_Timeline.png">
    <img src="/img/07_Misc_Auto-Rebuild_Timeline.png" alt="Screenshot of Timeline auto-rebuilding with new action items" />
</a>

(more info about this screenshot to come later...)

### Today's Agenda
<a href="/img/08_Todays_Agenda.png">
    <img src="/img/08_Todays_Agenda.png" alt="Screenshot of all open projects & action items shown at once" />
</a>

(more info about this screenshot to come later...)

### Forecasting, Task Selection, and Gantt Scrolling
<a href="/img/09_Forecasting_Task_Selection_Gantt_Scroll.png">
    <img src="/img/09_Forecasting_Task_Selection_Gantt_Scroll.png" alt="Screenshot of usability features when viewing the Timeline" />
</a>

(more info about this screenshot to come later...)

### Macros
<a href="/img/11_Macros.png">
    <img src="/img/11_Macros.png" alt="Screenshot of macros used for the Excel file" />
</a>