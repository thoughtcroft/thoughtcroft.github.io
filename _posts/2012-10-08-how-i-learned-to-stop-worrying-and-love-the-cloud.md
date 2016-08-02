---
layout:   post
title:    "How I learned to stop worrying and love the cloud"
date:     2012-10-08
category: cloud
tags:
  - books
  - cloud
  - ninefold
  - programming
  - ruby
---

*Originally published at Ninefold (2010-2015), a cloud services provider
I helped found.*

This is part 2 of the **Cloud programming** series.

When I first read the list of language features in the [Wikipedia entry
on the Ruby programming
language](https://en.wikipedia.org/wiki/Ruby_(programming_language), I
started screaming like a two year old.  Seriously, I did computer
science at ANU in the late (mumble)ties but I don’t remember learning
anything about first class continuations, closures, fibers or
duck-typing.  But as we learned in the Stanley Kubrick classic [Dr
Strangelove or: How I Learned to Stop Worrying and Love the
Bomb](http://www.imdb.com/title/tt0057012/), the best way to confront
your fear is to embrace it wholeheartedly.

In order to deliver on the challenge I set myself in [Part 1]({%
post_url 2012-04-16-a-social-cloud-experiment %}), I decided I’d better
give myself a crash-course in Ruby.

One of the first things I discovered about the Ruby programming world is
this: almost everyone uses a Mac (My theory? There are 57% more letters
in ‘Windows’ and brevity is one of the hallmarks of the Rubyist).  But
there is a perfectly adequate Ruby installer for Windows for those of us
still transitioning to the modern world (in fact my preferred Ruby
programming platform is now Ubuntu but let’s not get ahead of
ourselves).

I like good documentation, which can be a problem in open source
projects (The code is the documentation, dude! Really?), so I was
pleased to see that the Windows install includes [The Book of
Ruby](https://www.amazon.com.au/Book-Ruby-Hands--Guide-Adventurous-ebook/dp/B005EI84QA/ref=sr_1_4?s=digital-text&ie=UTF8&qid=1469689781&sr=1-4)
by Huw Collingbourne. However, the edition I got was based on Ruby 1.8
not the 1.9 version that was installed. So I headed over to [The
Pragmatic Bookshelf](https://pragprog.com/) for a copy of the complete
reference guide to [Programming Ruby
1.9](https://pragprog.com/book/ruby4/programming-ruby-1-9-2-0) by Dave
Thomas, otherwise known as the “PickAxe”. I have found this book
invaluable. On a side note, I love that you can set up your Pragmatic
account so that when you purchase an e-book it is automatically
delivered to your Dropbox and Kindle within minutes.

Next up, I installed an IDE (see, this Windows upbringing is hard to
shake), the excellent JetBrains RubyMine which is much more than an
editor.  I’ll explain more about this choice in a future blog.

I then embarked on a self-directed learning program.

I read through (and even understood some of) [Why’s Poignant Guide to
Ruby](http://poignant.guide/). This tries to twist your head into the
same dimension as the author which in my case was only moderately
successful, but it did whet my appetite with a taste of Ruby’s beauty.

[Try Ruby](http://tryruby.org/) is an interactive tutorial, providing a
basic introduction to the language via a browser based REPL (that’s
read-eval-print-loop or “console” for the rest of us). Like everyone
else, my first Ruby line of code was: `puts "Hello World!"`

[Learn Ruby the Hard Way](http://learnrubythehardway.org/) by Zed A.
Shaw and Rob Sobers is a PDF guide to learning programming in Ruby,
basically through repetition. There are 52 exercises that you read, type
in the sample code and run to check that you did it correctly. I
eventually found this quite tedious and skipped out at #42 but it did
get me started writing in the language.

Finally, I decided to climb onto the atomic bomb and ride that sucker
all the way down to ground zero: I went to [Rails
Camp](http://railscamps.com/)  11 which was held over a weekend in June
2012 at Koonjewarre, Springbrook in Queensland. This really was a
fantastic experience that I highly recommend and with the help of people
like Carl Woodward and Jeremy Grant, I managed to push out a working
Sinatra app to display a dashboard of scalable time-series charts of
Ninefold usage data using Highcharts.

From:

```ruby
puts "Hello World!"
```

To:

```ruby
@exclude ||= "AND t.account_id NOT IN (
  #{exclusions.map { |a| "'#{a}'" }.join(",") })" unless exclusions.nil?
```

I have evolved…
