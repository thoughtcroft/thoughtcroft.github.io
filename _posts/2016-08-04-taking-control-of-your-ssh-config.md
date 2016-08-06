---
layout:   post
title:    "Simplifying server access by using an SSH config file"
date:     2016-08-04
category: devops
tags:
  - devops
  - host
  - keys
  - password
  - security
  - ssh
---

If you are a devops engineer or a developer, at some point you will need
to jump onto a server to configure some software, access logs and
generally just check out what is happening. If you are like me, you'll
be doing it a lot.

You will likely be using ssh to login to the servers or at the very
least you will have enabled ssh key access for updating your source code
and pushing it to Github or Bitbucket. If you haven't, then you should
to improve security and make life easier.

It can be difficult to remember all those [ssh
options](http://linux.die.net/man/1/ssh). Well, if you aren't making the
[ssh config file](http://linux.die.net/man/5/ssh_config) do the hard
work for you, then you are doing it wrong. Here are some helpful tips.
(I'm going to assume you have basic knowledge about ssh keys, OK?)


### What is the role of the ssh config file?

In short, it is for supplying configuration to ssh. The ssh command
sources configuration parameters from the following sources in this order:

1. command-line options
1. the user's configuration file in `~/.ssh/config`
1. the system-wide configuration file in `/etc/ssh/ssh_config`
1. default values (these are listed in the system-wide config commented out)

Desired behaviour can be configured to always occur or specific
behaviour for a particular host. So for example rather than always
providing an IdentityFile (-i) option for a particular host,
you can instead add this to the config file and save yourself a ton of
typing.

### What does one look like?

Here's an extract from one that I use to make my life easier:

```bash
# specific host settings

Host github.com
  PreferredAuthentications publickey
  IdentityFile ~/.ssh/github_rsa

Host bitbucket.org
  PreferredAuthentications publickey
  IdentityFile ~/.ssh/bitbucket_rsa

# private CoreOS stuff

Host core-01
  HostName 172.17.8.101
Host core-02
  HostName 172.17.8.102
Host core-03
  HostName 172.17.8.103
Host core-??
  User core
  IdentityFile ~/.vagrant.d/insecure_private_key
  StrictHostKeyChecking no
  UserKnownHostsFile /dev/null

# general settings

IdentitiesOnly yes
ServerAliveInterval 60
```

Let's look at how this works in practice.

## Host name matching is the key

For each configuration parameter, the first obtained value is the one
that will be used. That's how command line options take precedence
over configuration file data. The same principle operates in the config
file.

Each *Host* section defines options that will take effect ifor any host that
matches the name referred to in the ssh command i.e.

    ssh user@host <options>

Patterns can include `*`, `?` and can be negated using `!`.
This matching starts at the top of the file and therefore config should
be entered with host specific values at the start and general values
towards the end.

For example, when I enter `ssh core-01` to log onto a vm running
CoreOS, it is equivalent to me typing:

```bash
ssh core@172.17.8.101 -i ~/.vagrant.d/insecure_private_key \
    -o StrictHostKeyChecking=no \
    -o UserKnownHostsFile=/dev/null \
    -o IdentitiesOnly=yes \
    -o ServerAliveInterval=60
```

### Use a different ssh key-pair per host

There are a lot of discussions like [this one on
StackEchange](http://security.stackexchange.com/questions/40050/what-is-the-best-practice-separate-ssh-key-per-host-and-user-vs-one-ssh-key-for)
regarding whether you should have separate private-public key pairs for
each server or one pair to rule them all.

One argument against using multiple keys is that the more additional
security you add, the more convenience you give up. With the ssh config
approach, it is no less convenient to use a per-host keypair than
one-pair for all hosts, apart from the act of generating the keys and
adding them to the config.

I use a different pair per host. Should there be some form of compromise
or some other reason that a given key-pair must be changed, there is
minimal effort to effect that for one host without any need to alter my
security with any other host or site.

Configuring each key in the ssh config makes this dead simple.

### Prevent 'known_hosts' problems when using virtual machines

If you are in the habit of spinning up and tearing down vms on a regular
basis, you will be aware of the problem whereby the ssh host key check
fails and prevents you from accessing the vm.

ssh maintains a list of all hosts that have been accessed in
*~/.ssh/known_hosts*. Whenever a host is accessed, it's previous entry
is checked and if different, access will be prevented to avoid
man-in-the-middle attacks. When vms are constantly being rebuilt, this
behaviour is very frustrating and requires you to remove the offending
key manually from *known_hosts*.

To avoid this, I use two settings:

* `UserKnownHostsFile=/dev/null`

  Throw away the host key instead of adding it to the *known_hosts*
  file. It isn't possible to suppress ssh trying to save the key but
  this has the same effect by saving it to /dev/null.

* `StrictHostKeyChecking=no`

  Prevent ssh from pausing to ask if the key should be saved the first
  time a new host is accessed.

In combination, these settings prevent interruptions when using
ssh to access vms that change frequently. It goes without saying that
these settings are **not recommended** for public or long-lived servers
as they protect you from potentially harmful attempts to compromise your
security.

### Other convenience settings

I'll leave it up to the reader to explore the other settings in detail,
but in short, here are the ones that I use to make life easier:

* `PreferredAuthentications=publickey`

  Don't fall back to other methods to access a server, just fail if
  public key access fails. If there is a networking problem
  or the public key has been deleted from *authorized_keys* on the host,
  then default behaviour is to fall back to asking for a userid /
  password. Since we know it is using an ssh key, don't bother doing
  this, just fail.

* `IdentitiesOnly=yes`

  Only use the identity files configured in the ssh config. This can
  avoid conflicts when there are other identities loaded via ssh-add.
  And there is no need to ssh-add any keys configured in the config
  file.

* `ServerAliveInterval=60`

  This sets the frequency in seconds for ssh to send KeepAlive messages
  to the host. If no response is received to these messages, the
  connection will be closed. In effect, this prevents ssh sessions with
  little activity from being timed out and closed.

I hope this is as helpful for you as it has been for me. If you have any
favourite tricks or settings, leave a comment below.
