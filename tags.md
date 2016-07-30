---
layout: page
title: Tags
permalink: /tags/
---

{% assign tags = site.tags | sort %}
{% assign max_tags = 0 %}
{% for tag in tags %}
  {% if tag.last.size > max_tags %} {% assign max_tags = tag.last.size %} {% endif %}
{% endfor %}

{% for tag in tags %}<span class="site-tag"><a href="/tag/{{ tag | first | slugify }}/" style="font-size: {{ tag | last | size | times: 100 | divided_by: max_tags }}%">{{ tag[0] | replace:'-', ' ' }} ({{ tag | last | size }})</a></span>&nbsp;
{% endfor %}
