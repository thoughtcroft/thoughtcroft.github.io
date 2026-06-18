require 'rubygems'
require 'rake'
require 'rdoc'
require 'date'

task :default => :serve

desc "Create new draft with optional args: title, cat, date"
task :post do
  title = ENV['title'] || "New Draft Post"
  slug = title.gsub(' ','-').downcase
  cat = ENV['cat'] || "random"
  date = ENV['date'] || Time.new.strftime('%Y-%m-%d')
  filename = "#{date}-#{slug}.md"
  dirname = '_drafts'
  path = File.join(dirname, filename)
  editor = ENV['EDITOR'] || "vim"

  post = <<-"EOF"
---
title:    "#{title}"
date:     #{date}
categories: [#{cat}]
tags:
  - #{cat}
---

# DRAFT!

EOF
  Dir.mkdir(dirname) unless Dir.exists?(dirname)
  File.open(path, 'w') { |f| f.puts post }
  exec "#{editor} #{path}"
end

desc "Generate blog files"
task :build, [:env] do |task, args|
  env = ( args.env || 'development' ).downcase
  system "JEKYLL_ENV=#{env} bundle exec jekyll build"
end

desc "Generate and serve locally"
task :serve do
  system "JEKYLL_ENV=development bundle exec jekyll serve --drafts"
end
