---
layout: post
title:  "Setting text edit mode on Access form fields made easy"
date:   2005-11-25
tags:   [access, control, form, textfield, vba]
---

The Locked and Enabled properties of text-based controls - combo boxes,
check boxes, list boxes and text boxes - control whether they can be
changed or entered. But for the life of me, I can never remember which
combination of true and false values gives me the look I am after. For
example, `Enabled=Yes` and `Locked=Yes` means the text field can be
entered but can't be changed.  Change this to `Enabled=No` and
`Locked=Yes` and you won't be able to enter or edit the field. Make it
`Enabled=Yes`, `Locked=No` and you can enter and edit and so on.

To make it easier on myself, I wrote this function to do the remembering
for me. As you can see, a lot of the work is done by the enumerated
constants definition - that's why I like to use them apart from the fact
that the VBA editor also reminds me what values I can use when I am
coding. I also like to use these constants as bitwise comparison flags
by making them different powers of 2 - easier to look at the example
than try and explain!

The net effect is that some sensible constant design simplifies coding
to two statements rather than a string of nested IFs.

{% gist thoughtcroft/f5850fba733eb0a2513cc5872166080b %}
