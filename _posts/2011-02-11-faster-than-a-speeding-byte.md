---
layout:   post
title:    "Faster than a speeding byte"
date:     2011-02-11
category: cloud
tags:
  - latency
  - network
  - ninefold
  - ping
---

> Originally published when I worked at Ninefold, an Australian cloud
> services provider I co-founded, that operated from 2010 - 2015.

There are a number of reasons why an Australian cloud server can be
better for you than an offshore equivalent. However, latency is easily
one of the most popular. So how much of a difference can it really make?

Latency is a synonym for delay, and in networking refers to a variety of
delays that can occur in transmitting data from one point to another.
These can include:

* propagation delays (mostly due to the speed of light and the distance
  the signal must travel)
* transmission delays (due to the physical properties of the
  transmission media); and
* processing delays (from the number of hops through different types of
  devices such as proxy servers and routers)

These delays can vary considerably due to traffic load or intermittent
faults but generally speaking, they increase the further the data has to
travel.

Typically, latency is measured by using ping tests which record the
round trip time for a packet of data to travel from source to
destination and back again, measured in milliseconds (ms). The longer
this takes, the slower a web app can display its web pages to the end
user and the customer experience suffers. Depending on how complex the
html, how many scripts must be parsed and how many images must be
retrieved, increased network latency measured in milliseconds can slow
down a webpage display by seconds.

For startups and developers targeting the Australian market, locally
based cloud is a much better alternative as it moves the data and
processing much, much closer to the customer.

To illustrate this point I ran 100 ping tests using Ping Plotter against
the Ninefold servers compared to a significant US based cloud provider.

These tests show that:

* The average latency for a typical Australian ADSL customer to Ninefold
  is 20 ms
* The average latency for the same customer to a US West Coast based
  cloud is 271 ms

That represents a speed advantage over 13 times in our favour.

When we consider the minimum times of 15ms to 266ms it becomes almost 18
times better latency locally.

Given the variations that different people will experience on the
interwebs, I feel entirely confident that we have at least a nine-fold
latency advantage over offshore cloud providers.

And that is something your customers will definitely appreciate.
