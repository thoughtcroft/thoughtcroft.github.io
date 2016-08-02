---
layout:   post
title:    "A social cloud programming experiment"
date:     2012-04-16
category: cloud
tags:
  - api
  - ninefold
  - programming
---

*Originally published at Ninefold (2010-2015), a cloud services provider
I helped found.*

This is part 3 of the **Cloud programming** series.

The cloud is a technological hub of remarkable innovation. However, an
idea has been developing in my mind for some time: a cloud programming
project where I document each step of the journey as I build something
unique and useful for our customers — something I hope you will join me
in creating.

I recently overheard a discussion that our support team had with a
customer that described the perfect candidate for such an experiment.
The customer is using Ninefold to provide the IT infrastructure for
their business, but many of the servers they use do not need to run
24x7. They would like to minimise their costs by turning off servers.
Customers with a single server in this position can control their costs
through our SimplePlan, but for customers with multiple servers a tool
to start or stop them in one operation would be just the ticket.

And that is what we will be building in this project: a simple tool to
start or stop all the cloud servers in a customer account in one
operation.

As one of the original founders of Ninefold, I have had an
intellectually challenging, exciting and at times scary two years of
intensive work with the cloud. Research, analysis, business case
development, project management, team building, branding, and software
development (draws breath) all played a vital role in the successful
launch of Ninefold.

I started my career as a developer, and while I have kept my hand in
various projects and still contribute to the Visual Basic for
Applications community — check out my VBA Adventures and the Microsoft
Project Holiday Import Wizard — I really don’t know a great deal about
the pointy end of programming in the cloud.

So I have decided to embark on a personal growth journey, and one that I
think will be very relevant to the Ninefold community: I am going to
learn a new programming language and build a stop/start server tool that
will be useful to our customers. And I hope that you will jump on board
and help me build it.

This idea has been germinating for some time. When we were researching
the cloud market it became clear that there were a number of **table
stake** characteristics necessary for Ninefold to be considered a true
cloud computing service:

* Self-service – Fully-automated provisioning under the customer’s
  control.
* On-demand – Resources can be spun up when required at any time (24/7).
* Scalable and elastic – Customers can increase and decrease their
  resources as required.
* Pay as you go – No contracts. Signup with a credit card and pay only
  for what you use.

Although not so obvious initially, there is another that underpins many
of the above:

* Cloud services must be exposed via an API

The Application Programming Interface (API) allows devs to control the
cloud through code. The Ninefold Portal uses our API to provision
resources, and our customers can also utilise a public subset to
automate provisioning from within their application. For example, a
server can monitor web-traffic and spin up additional servers in a
load-balancing group when the traffic crosses defined thresholds,
reducing the number of servers when the traffic falls below those
thresholds.

Our cloud provisioning API is described in detail in our customer wiki,
but we also recognise that most developers will prefer to abstract away
the differences between different cloud providers. So we have encouraged
and funded the inclusion of Ninefold in key cloud developer code
libraries such as:

* fog for Ruby
* Apache Libcloud for Python
* jclouds for Java and Clojure
* Ninefold on Nuget for .NET

This activity has only whet my appetite to become more hands-on in cloud
programming. In part two of this series I will share my choice of
language and the various tools that I will be employing on my project,
as well as how you can participate in the best traditions of open
source.

Keep an eye out for part two next week where we will flesh out the task,
commence the design, describe the language and tool choices and invite
your participation in the project. It will be fun and challenging and I
look forward to working with you. And believe me, given the learning
curve ahead I am depending on your help!
