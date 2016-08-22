require 'rubygems'
require 'rake'
require 'rdoc'
require 'date'
require 'shellwords'
require 'tmpdir'

task :default => :serve

desc "Create new draft with optional args: title, cat, date"
task :post do
  title = ENV['title'] || "New Draft Post"
  slug = title.gsub(' ','-').downcase
  cat = ENV['cat'] || "random"
  date = ENV['date'] || Time.new.strftime('%Y-%m-%d')
  filename = "#{date}-#{slug}.md"
  path = File.join('_drafts', filename)
  editor = ENV['EDITOR'] || "vim"

  post = <<-"EOF"
---
layout:   post
title:    "#{title}"
date:     #{date}
category: #{cat}
tags:
  - #{cat}
---

# DRAFT!

EOF
  File.open(path, 'w') { |f| f.puts post }
  exec "#{editor} #{path}"
end

desc "Generate blog files"
task :build, [:env] do |task, args|
  env = ( args.env || 'development' ).downcase
  system "JEKYLL_ENV=#{env} bundle exec jekyll build"
end

desc "Generate and publish blog to gh-pages"
task :publish do
  abort 'Please commit changes first!' if is_dirty?
  Rake::Task["build"].invoke("production")
  Dir.mktmpdir do |tmp|
    system "mv _site/* #{tmp}"
    system "git checkout -B master"
    system "rm -rf *"
    system "mv #{tmp}/* ."
    message = "Site updated at #{Time.now.utc}"
    system "git add ."
    system "git commit -am #{message.shellescape}"
    system "git push origin master --force"
    system "git checkout source"
  end
end

desc "Generate and serve locally"
task :serve do
  system "JEKYLL_ENV=development bundle exec jekyll serve --drafts"
end

def is_dirty?
  ! %x(git status -s).empty?
end
