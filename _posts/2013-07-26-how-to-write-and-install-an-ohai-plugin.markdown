---
layout: post
title:  "How to write and install an Ohai plugin, the Fight Club way"
date:   2013-07-26 11:43:39 +1000
tags:   [chef, cloud, infrastructure, ninefold, ohai, plugin, programming, ruby]
---

*Originally published at Ninefold (2010-2015), a cloud
services provider I helped found.*

This is part 4 of the **Cloud programming** series.

My loyal readers (Hi Mum and Dad!) will have noticed by now that we are
big fans of Chef at Ninefold. We have written recipes to configure and
manage our new multi-zone cloud server infrastructure. And we have also
built an amazing deployment framework utilising Private Chef to power
our new App Deploy service which we are finalising for launch in a
matter of weeks.

In my previous post - [Whipping Up Awesome with
Chef]({% post_url 2013-03-22-whipping-up-awesome-with-chef %}) - I provided a
brief explanation of the Chef universe and how the configuration state
of a node (server) is recorded on the Chef Server through attributes.
These attributes are defined by cookbook attribute files, recipes, roles
and environments and from data discovered about the node itself.  The
discovery process occurs at the start of every Chef client run and
automatically surfaces operating system platform details, memory and
processor usage, networking details, kernel data etc.

Departing from the cooking metaphor, the tool that discovers this
information is called Ohai. Apparently, Opscode can haz lotz da memez 4
mah codez!

Ohai discovery can be extended through plugins to provide custom
attributes as well. Plugins exist for AWS and other cloud providers and
so I decided to write one to expose Ninefold meta-data about servers
such as:

* availability zone
* service offering
* public ip address

As I grappled with the Ohai documentation I
was painfully reminded of my first experience of Chef: *there's something
someone isn't telling me because it's not making sense*. In the case of Ohai,
not only is the plugin authoring scant on detail but the method of
getting the plugin installed is subject to the 1st and 2nd Rules of Fight Club:

1. you do not talk about how to install Ohai plugins
1. you **DO NOT** talk about how to install Ohai plugins

But never fear, Wazza is here to take you through it, one step at a time
(that's the 5th Rule of Fight Club).

### Step 1 – create a new cookbook

Here's one I created earlier: [chef-ninefold-ohai](https://github.com/ninefold/chef-ninefold-ohai).
Feel free to fork this for your own use. The key things here are the plugin file itself in
`/files/default/plugins/ninefold.rb` and the default recipe which will do
the installation on the node `/recipes/default.rb`

### Step 2 – write the plugin

Essentially, you are going to write some simple ruby to expose
attributes as a Mash (this is a built-in Chef class which provides a
Hash with indifferent access i.e. you can access attributes as
params[:key] or params['key'], params[:key][:subkey] or
params.key.subkey etc.  So given a plugin called ninefold, we simply
need:

```ruby
provides 'ninefold'
ninefold Mash.new
# obtain meta_data from CloudPlatform
ninefold['availability-zone'] = meta_data[:zone]
# example only
```

Check the cookbook repo to see the actual meta-data retrieval code and
attribute setting
[here](https://github.com/ninefold/chef-ninefold-ohai/blob/master/files/default/plugins/ninefold.rb).

### Step 3 – tell Chef how to install the plugin

The simplest way to install the plugin is by using Chef itself.  Create
a default.rb recipe in the cookbook.

The first trick is to tell Chef to drop the plugin file onto the node
where Ohai can find it.  We do this using the 'ohai' cookbook.  Simply
create a key under the node's 'ohai.plugins' attribute which is named
after our cookbook (ninefold_ohai) with the name of the file folder in
that cookbook where the plugin can be found (plugins).  The ohai default
recipe will do the hard work for us by copying the ninefold.rb file into
the custom plugins directory on the node.  Any other plugin files in
this cookbook will also be installed, so this is an easily extensible
way of installing multiple plugins.

```ruby
node.set['ohai']['plugins']['ninefold_ohai'] = 'plugins'
include_recipe 'ohai'
```

The second trick is to update the chef-client configuration with the
custom plugin path so that Ohai will load and evaluate the plugin at the
start of each chef-client convergence.  This is very simple using the
chef-client cookbook which will ensure that the client.rb file is
properly configured.

```ruby
include_recipe 'chef-client::config'
```

Don't forget to tell Chef that you need those two cookbooks whenever the
ninefold_ohai cookbook is being loaded.  In the metadata.rb file add

```ruby
depends "ohai"
depends "chef-client"
```

### Step 4 – ensure the plugin is loaded at the start of the node convergence

Simply add the following as the first item in the node's run_list -
'recipe[ninefold_ohai]' – and Ohai will automatically load our plugin
and populate the 'ninefold' attributes for use by subsequently loaded
recipes, or for chef searches after the node has converged.

> Tip: if it isn't first in the run_list, recipes loaded before ninefold_ohai won't be able to access the new custom attributes.

Now remember the 8th Rule of Fight Club: *if this is the first time you
have looked at Ohai custom plugins, you HAVE to write an Ohai custom
plugin*.

With this guide, it should be a piece of cake (now there's a cooking
term crying out for an Opscode Chef feature).
