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

input_root = os.getcwd()
output_root = os.path.join(os.getcwd(), 'html')

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

force = False
eval_globals = dict([(fn, getattr(math, fn)) for fn in dir(math)])

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


def process(template, data, meta):
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
                source = data[match.group(2)] if match.group(2) != '*' else meta['data']
                template = match.group(3)
                print (source)
                print (template)
                if isinstance(data[match.group(2)], list):
                    return '\n'.join([str(eval(f"f'{template}'", data, item)) for item in source])
                if isinstance(data[match.group(2)], dict):
                    return '\n'.join([str(eval(f"f'{template}'", {**data, **{'key': key}}, source[key])) for key in source])
            item = str(match.group(1))
            if item.startswith('#'):
                processed = str(eval(item[1:], context, data))
                return markdown2.markdown(processed, extras=extras)
            return str(eval(item, context, data))
        except:
            return match.group(0)

    return pattern.sub(replace, template)

def fix_name(name):
    translation = str.maketrans(' /()?$%,-#.&:', '_____________')
    name = name.translate(translation)
    if re.match('[0-9]', name):
        name = '_' + name
    return name.strip()


def process_dir(path):
    '''Recursively process all the files in the directory.'''
    log (f'Processing directory {path}')
    fullpath = os.path.join(input_root, path)

    # Don't process the output directory if it's in the tree
    if os.path.normpath(fullpath) == os.path.normpath(output_root):
        return

    #print (os.path.normpath(fullpath))
    #print (os.path.normpath(output_root))
    output_path = os.path.join(output_root, path)
    if not os.path.exists(output_path):
        os.makedirs(output_path)
        log(f'make dir {output_path}')

    filenames = [filename for filename in os.listdir(fullpath)]
    filenames.sort(key=lambda x: os.path.getmtime(os.path.join(fullpath, x)))
    directories = [file for file in filenames if os.path.isdir(os.path.join(fullpath, file))]
    image_files = [file for file in filenames if os.path.splitext(file)[1].lower in ['.png', '.jpg']]
    content = [os.path.splitext(filename) for filename in filenames]
    yaml_files = [file for file in content if file[1] == '.yaml']
    csv_files = [file for file in content if file[1] == '.csv']
    markdown_files = [file for file in content if file[1] == '.md']


    # Read data from csv and yaml files

    data = {'index': {'template': 'index.html', 'modification_time': datetime.timestamp(datetime.now())}}
    # Using the first row as the property names, generate a single entry for each subsequent row of the csv
    for name, extension in csv_files:
        fullname = os.path.join(fullpath, f'{name}{extension}')
        mtime = os.path.getmtime(fullname)
        with open(fullname) as csv_file:
            header = None
            for row in csv.reader(csv_file):
                if header:
                    item_name = row[0].replace('/', '_').replace(':', '_')
                    if item_name not in data:
                        data[item_name] = { 'template':  os.path.join(fullpath, f'{name}.html') }
                    data[item_name]['modification_time'] = mtime
                    for i in range(len(header)):    # Allow for Jira's habbit of repeating header for list fields
                        if header[i] in data[item_name]:
                            data[item_name][header[i]] += f'\n\n{row[i]}'
                        else:
                            data[item_name][header[i]] = row[i]
                else:
                    header = [fix_name(item) for item in row]

    # Process YAML files overwriting data from csv if item names match
    for name, extension in yaml_files:
        fullname = os.path.join(fullpath, f'{name}{extension}')
        with open(fullname) as stream:
            try:
                yaml_data = yaml.safe_load(stream)
                if name not in data:
                    data[name] = dict([(fix_name(key), yaml_data[key]) for key in yaml_data])
                else:
                    for key in yaml_data:
                        data[name][fix_name(key)] = yaml_data[key]
                if 'Title' not in data[name]:
                    data[name]['Title'] = name
                data[name]['modification_time'] = os.path.getmtime(fullname)
            except yaml.YAMLError as exc:
                print (exc)

    templated = [templated for templated in data if 'template' in data[templated]]

    for name in templated:
        if not os.path.exists(data[name]['template']):
            data[name]['template'] = os.path.join(fullpath, data[name]['template'])

    found = [name for name in templated if os.path.exists(data[name]['template'])]
    missing = [name for name in templated if name != 'index' and not os.path.exists(data[name]['template'])]
    if len(missing) > 0:
        log (f'Template file not found for {", ".join(missing)}')

    found.sort()
    templates = set([data[name]['template'] for name in found])
 
    index = [f' - [{name}]({name}.html)' for name in found if name != 'index']
    meta = {'index': '\n'.join(index), 'data': data}

    # Generate files from template files

    for template_file in templates:
        with open(template_file) as file:
            template  = file.read()
        for key in data:
            instance = data[key]
            if instance.get('template', '') == template_file:
                output_file = os.path.join(output_path, f'{key}.html')
                if force or not os.path.exists(output_file) \
                    or os.path.getmtime(output_file) < data[key]['modification_time'] \
                    or os.path.getmtime(output_file) < os.path.getmtime(template_file):
                    mdate = date.fromtimestamp(data[key]['modification_time']).strftime('%d/%m/%Y')
                    source = process(template, instance, meta)
                    if os.path.splitext(template_file)[1] == '.md':
                        content = markdown2.markdown(source, extras=extras)
                        result = boilerplate.replace('%TITLE%', key)\
                            .replace('%UPDATED%', mdate)\
                            .replace('%CONTENT%', content)
                    else:
                        result = source.replace('%TITLE%', key)\
                            .replace('%UPDATED%', mdate)
                    output_file = os.path.join(output_path, f'{key}.html')
                    log (f'Updating {output_file}')
                    with open(output_file, 'w') as file:
                        file.write(result)

    for name, extension in markdown_files:
        fullname = os.path.join(fullpath, f'{name}{extension}')
        output_file = os.path.join(output_path, f'{name}.html')
        if force or not os.path.exists(output_file) \
            or os.path.getmtime(output_file) < os.path.getmtime(fullname):
            mdate = date.fromtimestamp(os.path.getmtime(fullname)).strftime('%d/%m/%Y')
            with open(fullname) as file:
                markdown  = file.read()
                source = process(markdown, data, meta)
                content = markdown2.markdown(source, extras=extras)
                result = boilerplate.replace('%TITLE%', name)\
                .replace('%UPDATED%', mdate)\
                .replace('%CONTENT%', content)
                log (f'Updating {output_file}')
                with open(output_file, 'w') as file:
                    file.write(result)

    for image_file in image_files:
        fullname = os.path.join(fullpath, image_file)
        destination = os.path.join(output_path, image_file)
        log (f'Copying {fullname}')
        if force or not os.path.exists(destination) or os.path.getmtime(destination) < os.path.getmtime(fullname):
            shutil.copy(fullname, destination)

    for directory in directories:
        process_dir(os.path.join(path, directory))

process_dir('')
