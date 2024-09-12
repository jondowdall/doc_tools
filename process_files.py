#!/usr/bin/python3

import markdown2
import os
import re
import sys
import yaml


boilerplate = '''
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>%TITLE%</title>
    <link rel="stylesheet" href="style.css">
  </head>
  <body>
    %CONTENT%
    <script src="index.js"></script>
  </body>
</html>
'''

extras = ['tables', 'strike']


def process(template, data, meta):
    pattern = re.compile('\{(\S[^\}]*)\}')
    unroll = re.compile('\.\.\.([^:]+):(.*)')
    def replace(match):
        item = match.group(1)
        loop = unroll.match(item)
        if loop and loop.group(1) in data:
            result = []
            for item in data[loop.group(1)]:
                line = loop.group(2)
                for key in item:
                    line = line.replace(key, str(item[key]))
                result.append(line)
            return '\n'.join(result)

        if item in data:
            return str(data[item])
        if item.strip('%') in meta:
            return meta[item.strip('%')]
        return match.group(0)

    return pattern.sub(replace, template)


def process_dir(path):
    '''Recursively process all the files in the directory.'''

    content = [os.path.splitext(filename) for filename in os.listdir(path)]
    yaml_files = [file for file in content if file[1] == '.yaml']
    markdown_files = [file for file in content if file[1] == '.md']
    data = {}
    for name, extension in yaml_files:
        with open(f'{name}{extension}') as stream:
            try:
                data[name] = yaml.safe_load(stream)
            except yaml.YAMLError as exc:
                print (exc)

    keys = [key for key in data
            if 'template' in data[key] and os.path.exists(data[key]['template'])]

    keys.sort()
    templates = set([data[key]['template'] for key in keys])

    index = [f' - [{key}]({key}.html)' for key in keys]
    meta = {'index': '\n'.join(index)}

    for template_file in templates:
        with open(template_file) as file:
            template  = file.read()
        for key in data:
            instance = data[key]
            if instance.get('template', '') == template_file:
                source = process(template, instance, meta)
                content = markdown2.markdown(source, extras=extras)
                result = boilerplate.replace('%TITLE%', key).replace('%CONTENT%', content)
                with open(f'{key}.html', 'w') as file:
                    file.write(result)

    for name, extension in markdown_files:
        with open(f'{name}{extension}') as file:
            markdown  = file.read()
            source = process(markdown, data, meta)
            content = markdown2.markdown(source, extras=extras)
            result = boilerplate.replace('%TITLE%', name).replace('%CONTENT%', content)
            with open(f'{name}.html', 'w') as file:
                file.write(result)


#with open("sandbox/test.yaml") as stream:
#     try:
#         print (yaml.safe_load(stream))
#     except yaml.YAMLError as exc:
#         print (exc)

process_dir(os.getcwd())
