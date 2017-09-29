# -*- encoding: utf-8 -*-

lib = File.expand_path('../lib/', __FILE__)
$:.unshift lib unless $:.include?(lib)
 
require 'goog/version'

Gem::Specification.new do |s|
  s.name = "goog"
  s.version = Goog::VERSION

  s.platform    = Gem::Platform::RUBY
  s.authors     = ["Tim Harrison"]
  s.email       = ["heedspin@gmail.com"]
  s.homepage    = "http://github.com/heedspin/goog"
  s.summary     = "Simple google api wrappers to help my code!"
  s.description = "Code that's shared between many of my projects"
 
  s.required_rubygems_version = ">= 2.0.0" 
  s.add_dependency 'plutolib'
  s.files        = Dir.glob("{bin,lib}/**/*") + %w(LICENSE README.md)
  s.require_path = 'lib'
end

