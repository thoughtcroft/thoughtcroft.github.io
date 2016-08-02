---
layout: post
title:  "Finding the name of a control in an Access subform"
date:   2005-11-25
tags:   [access, form, subform, vba]
---

I'm currently developing a Microsoft Access based system and found
myself continually needing to work out how to access a property of the
control that contains a subform from code running in the subform (for
example to get at the Tag property).

This function will do that by walking the controls collection of the
subform's parent form looking for any subforms and then compares the
hWnd (basically Windows internal "handle" for that window) of that
control with the hWnd of our subform.

Once found, we construct the name of the control using the appropriate
name format. If we want to use the name in code then the short format is
fine but if it is to be used in a query then we need the long version
which may necessitate walking up the hierarchy if in fact the parent
form is itself a subform (forms can be nested to three levels). This is
achieved by calling the function recursively on the parent.


{% gist thoughtcroft/f3f541bb864031a79c9091a049fe6eb2 %}
