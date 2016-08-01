---
layout: post
title:  "Whipping up awesome with Chef"
date:   2013-03-22
tags:   [automation, chef, cloud, devops, infrastructure, ninefold]
---

*Originally published at Ninefold (2010-2015), a cloud
services provider I helped found.*

This is part 3 of the **Cloud programming** series.

When I started this blog series, I had a project in mind which required
me to learn Ruby and how to use the Fog cloud abstraction library so
that I could start and stop groups of Ninefold virtual servers at will.
That project is still on the drawing board :), for now I have moved on to
something far more important to our customers.

On my rubyist reinvention journey, I kept coming across something called
Chef, an open source configuration management tool managed by OpsCode
out of Seattle.  And as I listened to people's enthusiasm about the
power of Chef, I realised that this was the answer to two of our biggest
challenges:

* managing Ninefold's expanding infrastructure without expanding our ops
  team at the same rate;
* providing simple to use app deployment for our customers that also
  allows them to leverage the full power of our production grade
  multi-zone infrastructure.

Here's the quick explanation of Chef: **your infrastructure as code**.

OK, a bit cryptic, so try this longer definition:

1. Chef is a systems integration framework, designed to bring configuration
management to your entire infrastructure. It is an open source project
with a vibrant community contributing to the base software and sharing
cookbooks (abstracted definitions of resource configuration).

1. Servers are known as **Nodes**. Nodes have attributes that describe the
current or desired state of its configuration and provide one of several
mechanisms for communicating config changes.

1. **Cookbooks** are collections of files which configure the nodes to a
desired end-state. **Recipes** are written using ruby and a recipe DSL which
is readily extensible.

1. The various elements of the configuration are known as **resources** e.g.
apache web server, iptables firewall, mysql database etc.  Recipes are (should be)
[idempotent](https://en.wikipedia.org/wiki/Idempotence#Computer_science_meaning)
i.e. re-running the recipe on a node multiple times always returns the system to
an identical state.

In terms of software, chef-client runs on the node and talks to a server
which holds the cookbooks and attribute status of all registered nodes.
There is a special chef-solo version of the server which runs on the
node, an open source chef-server, an OpsCode Hosted Chef server as a
service and a licensed, multi-tenant, HA edition called Private Chef.

A chef-client run operates in two stages. During compilation, the
various code files – libraries, attributes, definitions, recipes etc –
are loaded and evaluated. A resource list is built from the DSL in each
recipe in the order they appear. During convergence, each resource is
configured according to the relevant DSL configuration.
Rather than go into a lot more detail about Chef – and believe me, there
is a lot of detail to get your head around – I am going to share my Top
Four Tips for Chef.

### Tip 1 - Understand the difference in Compilation vs Convergence

During compilation, any normal Ruby code is evaluated when the relevant
file – attribute, library, recipe etc – is first loaded.  If you want to
calculate some value after a resource has been converged and then save
it in an attribute, you will need to delay the evaluation by wrapping
the relevant code in a ruby_block resource:

```ruby
ruby_block “save node state” do
  block do
    node.set['some_host']['configured_time'] = Time.now.to_s
    node.save
  end
  action :create
end
```

For a more detailed and highly readable explanation, check out the
[Anatomy of a Chef Run](https://docs.chef.io/chef_client.html).

### Tip 2 -  The cause of chef run terminating with "undefined method '[]' for nil:NilClass"

This initially perplexing message is almost always due to trying to
reference a node attribute that doesn't exist. For example if I am
expecting to use `node['my_cookbook']['app']['version']` somewhere in my
recipe, I will get the value 'nil' if `node['my_cookbook']['app']` exists
but there is no `['version']` attribute.  But if
`node['my_cookbook']['app']` also doesn't exist then I get the dreaded
undefined method '[]' since I am in effect trying to reference
`nil['version']` and nil doesn't have a '[]' method.

To avoid this common occurrence in your early recipe writing, ensure
attributes exist before you use them.  Two methods are:

```ruby
if node.attribute?('version')
  # there is some attribute somewhere called 'version'
  # perhaps not the best method
end

if node['my_cookbook']['app'] && node['my_cookbook']['app']['version']
  # there is an 'app' and a 'version' for 'app'
end
```

### Tip 3 -  Essential knife plug-in: knife-block

Knife is the command line tool for managing your chef server and nodes
from your workstation. If you use Chef Server it is likely that you will
have more than one of these.  At Ninefold, we use Private Chef and as
well as multiple instances of the server, we also have multiple
organisations within each server – you can think of each organisation as
a logical Chef Server with separate cookbooks etc but sharing the one
client-validation key.

Managing multiple knife.rb configuration files is a bit of a nightmare,
and requires you to place the relevant knife.rb, chef-validator.pem and
client.pem files into each project. But then if you want to push a
cookbook that you are working on from your development chef to your
testing chef, there is a spot of juggling required. That is until you
install the
[knife-block plugin](https://github.com/knife-block/knife-block)
by Green and Secure IT Limited.

Place all your knife.rb and .pem files into your /home/.chef/ directory
and rename each knife.rb file as knife-{something}.rb e.g. knife-dev.rb,
knife-test.rb, knife-wazza-is-awesome.rb. `knife block dev` will create
a knife.rb symlink to the knife-dev.rb.

knife walks up the directory structure looking in ../.chef/ to find the
knife.rb configuration so placing all those files in /home/.chef/ means
knife will always find the configuration required for your current chef
context.  At any time you can find out your choices by

```bash
$ knife block list
  The available chef servers are:
    * dev
    * test
    * wazza-is-awesome [ Currently Selected ]
```

And switch context simply using

```bash
$ knife block test
  The knife configuration has been updated to use test

$ knife block list
  The available chef servers are:
    * dev
    * test [ Currently Selected ]
    * wazza-is-awesome
```

### Tip 4 -  Manage cookbook dependencies using Berkshelf

As [Berkshelf](http://berkshelf.com) states “If you're familiar with Bundler, then
Berkshelf is a breeze”.

A cookbook's metadata.rb file uses 'depends' clauses to specify cookbook
dependencies – at Ninefold, we almost always specify exact versions to
isolate us from potential breaking changes. Chef ensures that these
versions are loaded onto the node at the start of a run provided they
are present on the Chef Server. How do they get on to the server in the
first place? You upload them using knife.  But if you have a number of
cookbook development projects and if you are reliant on specific branch
versions of community cookbooks, then managing this process is very
difficult. Until Berkshelf.

And ermagherd, Berkshelf ers ersum!

Cookbooks can be easily managed in single repositories. Dependency data
can be drawn from the cookbook's metadata.rb and the source of dependent
cookbooks can be defined in various ways. Cookbooks are installed into
/home/.berkshelf and this is where they are sourced by default but if
missing from there they can be sourced from the Opscode Community
Cookbooks site, a specific branch of a git repository, a path on the
workstation or from a chef server.  Once the cookbooks have been
installed or updated they can be bulk uploaded into the Chef Server.

Ninefold's cookbook development and customer provisioning process makes
extensive use of Berkshelf and I highly recommend reading an
introduction to authoring cookbooks by Jamie Winsor, the creator of
Berkshelf.
