from argtools import command, argument
import csv
from datetime import date
from datetime import datetime
import difflib
import markdown2
import math
import os
import re
import shutil
import yaml


boilerplate = '''
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>%TITLE%</title>
    <link rel="stylesheet" href="style.css">
    <style>
        html {
            font-family: sans-serif;
            box-sizing: border-box;
        }

        *,
        *:before,
        *:after {
            box-sizing: inherit;
        }

        table, th, td {
            border-collapse: collapse;
            border: 1px solid gray;
        }

        th, td {
            padding: 0.3em;
        }

        th {
            background: lightgray;
            font-weight: bold;
            position: sticky;
            top: 0;
        }

        .updated {
            position: fixed;
            top: 0.5em;
            right: 0.5em;
            font-weight: bold;
        }

        .mermaid-pre {
            visibility: hidden;
        }
    </style>
</head>
<body>
    <div class="updated">Updated: %UPDATED%</div>
    %CONTENT%
    <script src="script.js"></script>
    <script type="module" defer>
        import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@9/dist/mermaid.esm.min.mjs';
        mermaid.initialize({
            securityLevel: 'loose',
            startOnLoad: true
        });
        let observer = new MutationObserver(mutations => {
            for(let mutation of mutations) {
            mutation.target.style.visibility = "visible";
            }
        });
        document.querySelectorAll("pre.mermaid-pre div.mermaid").forEach(item => {
            observer.observe(item, { 
            attributes: true, 
            attributeFilter: ['data-processed'] });
        });
    </script>
</body>
</html>
'''

eval_globals = dict([(fn, getattr(math, fn)) for fn in dir(math) if not fn.startswith('_')])

extras = ['tables', 'strike', 'cuddled-lists', 'fenced-code-blocks',
          'header-ids', 'numbering', 'task-list', 'wiki-tables', 'mermaid']

def log(text, level=0):
    '''Print log data'''
    print(text)

def sum(expr):
    result = 0
    for item in list:
        result += eval(expr, data, item)
    return result

def to_dict(obj):
    if isinstance(obj, dict):
        return obj
    elif isinstance(obj, list):
        return dict([(f'_{i}', obj[i]) for i in range(len(obj))])
    return {'key': obj}

def process(template, item, meta):
    pattern = re.compile('\{(\S[^\}]*)\}|\[([^:]+):\s*([^\]]+)\]')
    context = {**eval_globals, **meta}

    def sum(list, expr):
        result = 0
        for item in list:
            result += eval(expr, data, item)
        return result
    
    context['sum'] = sum

    def replace(match):
        try:
            if match.group(2):
                if match.group(2) == '*':
                    source = meta['data']
                elif match.group(2) == '**':
                    source = item
                elif match.group(2).startswith('*'):
                    source = meta[match.group(2)[1:]]
                else:
                    source = item[match.group(2)]
                template = match.group(3)
                if isinstance(source, list) or isinstance(source, set):
                    return '\n'.join([str(eval(f"f'{template}'", meta['data'], to_dict(item))) for item in source])
                if isinstance(source, dict):
                    return '\n'.join([str(eval(f"f'{template}'", {**meta['data'], **{'key': key}}, item)) for key in source])
            element = str(match.group(1))
            if element.startswith('#'):
                processed = str(eval(element[1:], context, item))
                return markdown2.markdown(processed, extras=extras)
            return str(eval(element, context, item))
        except Exception as e:
            print(e)
            return match.group(0)
    return pattern.sub(replace, template)


def fix_name(name):
    translation = str.maketrans(' /()?$%,-#.&:', '_____________')
    name = name.translate(translation)
    if re.match('[0-9]', name):
        name = '_' + name
    return name.strip()


def process_dir(source, destination, force):
    '''Recursively process all the files in the directory.'''

    source_dir = os.path.abspath(source)
    destination_dir = os.path.abspath(destination)
    log (f'Processing directory {source_dir}')

    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)
        log(f'make dir {destination_dir}')

    filenames = [filename for filename in os.listdir(source_dir)]
    filenames.sort(key=lambda x: os.path.getmtime(os.path.join(source_dir, x)))
    directories = [file for file in filenames if os.path.isdir(os.path.join(source_dir, file))]
    image_files = [file for file in filenames if os.path.splitext(file)[1].lower() in ['.png', '.jpg']]

    content = [os.path.splitext(filename) for filename in filenames]
    yaml_files = [file for file in content if file[1] == '.yaml']
    csv_files = [file for file in content if file[1] == '.csv']
    template_files = dict([(os.path.join(source_dir, name), set()) for name in filenames
                            if os.path.splitext(name)[1].lower() in ['.md', '.html']])
    markdown_files = [file for file in content if file[1] == '.md']

    tree = {}
    for directory in directories:
        if os.path.join(source_dir, directory) != destination_dir:
           tree[directory] = process_dir(os.path.join(source_dir, directory), os.path.join(destination_dir, directory), force)
    
    # Read data from csv and yaml files
    data = {}
    # Use the first row as the property names, generate a single entry for each subsequent row of the csv
    for filename, extension in csv_files:
        fullname = os.path.join(source_dir, f'{filename}{extension}')
        mtime = os.path.getmtime(fullname)
        with open(fullname) as csv_file:
            header = None
            for row in csv.reader(csv_file):
                if header:
                    name = row[0].replace('/', '_').replace(':', '_')
                    if name not in data:
                        data[name] = { 'template': os.path.join(source_dir, f'{filename}.html'), 'content': {} }
                    data[name]['modification_time'] = mtime
                    content = data[name]['content']
                    for i in range(len(header)):    # Allow for Jira's habbit of repeating header for list fields
                        if header[i] in content:
                            content[header[i]] += f'\n\n{row[i]}'
                        else:
                            content[header[i]] = row[i]
                else:
                    header = [fix_name(item) for item in row]

    # Process YAML files overwriting data from csv if item names match
    for name, extension in yaml_files:
        fullname = os.path.join(source_dir, f'{name}{extension}')
        with open(fullname) as stream:
            try:
                yaml_data = yaml.safe_load(stream)
                if name not in data:
                    data[name] = { 'content': dict([(fix_name(key), yaml_data[key]) for key in yaml_data]) }
                    if 'template' in yaml_data:
                        data[name]['template'] = yaml_data['template']
                else:
                    for key in yaml_data:
                        data[name]['content'][fix_name(key)] = yaml_data[key]
                if 'Title' not in yaml_data:
                    data[name]['Title'] = name
                data[name]['modification_time'] = os.path.getmtime(fullname)
            except yaml.YAMLError as exc:
                print (exc)

    for name, item in data.items():
        if 'template' in item:
            template = os.path.normpath(item['template'])
            if not os.path.exists(template):
                template = os.path.normpath(os.path.join(source_dir, item['template']))
            if os.path.exists(template):
                if template in template_files:
                    template_files[template].add(name)
                else:
                    template_files[template] = {name} 
            else:
                log(f'Template ({template}) for {name} not found!')
        elif name in [os.path.splitext(os.path.basename(filename))[0] for filename in template_files]:
            template = [template for template in template_files if name == os.path.splitext(os.path.basename(template))[0]][0]
            template_files[template].add(name)
        else:
           log(f'no template found for {name}')

    meta = {
        'directories': directories,
        'tree': tree,
        'images': image_files,
        'templates': template_files,
        'data': data
    }

    for template_file, names in template_files.items():
        with open(template_file) as file:
            template  = file.read()
        
        if len(names) > 0:
            sources = dict([(name, data[name]) for name in names])
        else:
            mtime = os.path.getmtime(template_file)
            name = os.path.splitext(os.path.basename(template_file))[0]
            sources = {name: { 'modification_time': mtime, 'content': data }}
    
        for name, item in sources.items():
            output_file = os.path.join(destination_dir, f'{name}.html')
            if force or not os.path.exists(output_file) \
                or os.path.getmtime(output_file) < item['modification_time'] \
                or os.path.getmtime(output_file) < os.path.getmtime(template_file):
                mdate = date.fromtimestamp(item['modification_time']).strftime('%d/%m/%Y')
                source = process(template, item['content'], meta)
                if os.path.splitext(template_file)[1] == '.md':
                    content = markdown2.markdown(source, extras=extras)
                    result = boilerplate.replace('%TITLE%', name)\
                        .replace('%UPDATED%', mdate)\
                        .replace('%CONTENT%', content)
                else:
                    result = source.replace('%TITLE%', name)\
                        .replace('%UPDATED%', mdate)
                log (f'Updating {output_file}')
                with open(output_file, 'w') as file:
                    file.write(result)

    for image_file in image_files:
        fullname = os.path.join(source_dir, image_file)
        destination = os.path.join(destination_dir, filename)
        log (f'Copying {fullname}')
        if force or not os.path.exists(destination) or os.path.getmtime(destination) < os.path.getmtime(fullname):
            shutil.copy(fullname, destination)

    return tree

@command
@argument('--source', default=os.getcwd(), help='directory of source files')
@argument('--destination', default=os.path.join(os.getcwd(), 'html'), help='destination to write files to')
@argument('--force', action='store_true', help='an optional argument')
def main(args):
    """ One line description here

    Write details here (printed with --help|-h)
    """
    process_dir(args.source, args.destination, args.force)


if __name__ == '__main__':
    command.run()
