require 'rubygems'
require 'rake'
require 'rdoc'
require 'date'

task :default => :serve

desc "Create new draft, e.g. rake post[\"My Title\",design,2026-06-19]"
task :post, [:title, :cat, :date] do |task, args|
  title = args.title || "New Draft Post"
  slug = title.gsub(' ','-').downcase
  cat = args.cat || "random"
  date = args.date || Time.new.strftime('%Y-%m-%d')
  filename = "#{date}-#{slug}.md"
  dirname = '_drafts'
  path = File.join(dirname, filename)
  editor = ENV['EDITOR']

  post = <<-"EOF"
---
layout: post
title: #{title}
date: #{date}
description:
categories:
- #{cat}
tags:
- #{cat}
---

# DRAFT!

EOF
  Dir.mkdir(dirname) unless Dir.exist?(dirname)
  File.open(path, 'w') { |f| f.puts post }
  system("#{editor} #{path} &")
end

desc "Generate blog files"
task :build, [:env] do |task, args|
  env = ( args.env || 'development' ).downcase
  system("JEKYLL_ENV=#{env} bundle exec jekyll build")
end

desc "Generate and serve locally"
task :serve do
  system("JEKYLL_ENV=development bundle exec jekyll serve --drafts")
end
