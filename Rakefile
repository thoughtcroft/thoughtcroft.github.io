require 'rubygems'
require 'rake'
require 'rdoc'
require 'date'
require 'tmpdir'

task :default => :serve

desc "Create new blog post args: title, date (optional)"
task :post do
  title = ENV['title'] || "New Post"
  slug = title.gsub(' ','-').downcase
  date = ENV['date'] || Time.new.strftime('%Y-%m-%d')
  filename = "#{date}-#{slug}.md"
  path = File.join('_posts', filename)
  editor = ENV['EDITOR'] || "vim"

  post = <<-"EOF"
---
layout: post
title:  "#{title}"
date:   #{date}
tags:   []
---

EOF
  File.open(path, 'w') { |f| f.puts post }
  exec "#{editor} #{path}"
end

desc "Generate blog files"
task :generate, [:env] do |task, args|
  env = args.env || 'development'
  system "JEKYLL_ENV=#{env.downcase} bundle exec jekyll build"
end

desc "Generate and publish blog to gh-pages"
task :publish do
  Rake::Task["generate"].invoke("production")
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
  system "bundle exec jekyll serve"
end
