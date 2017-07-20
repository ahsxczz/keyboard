#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import unicode_literals
import os.path
import traceback
from datetime import datetime
from path import Path
from xlrd import open_workbook, xldate_as_tuple
from dateutil.parser import parse
import emails
from emails.template import JinjaTemplate as T
import click

click.disable_unicode_literals_warning = True

folder_help = 'Folder or file path of Excel, default: H:\\TE\\Proyekto\\configs'
level_help = 'Alert level, default: 3'
level_maps_help = 'maps of level: days, default: 1:180,2:90,3:30'
template_help = 'jinja2 template for rendering, default: templates\one_month.html'
host_help = 'email server, default: shateweb'
sender_help = 'Sender address, name and email separated by $, default: \
Validation Expiring Alert$validation_alert@jabil.com'
date_format_help = 'date format, default: %d-%b-%Y'
date_columns_help = 'other columns (zero-index) contain date, default: 8,9'
comment_columns_help='other columns (zero-index) contain comment, default: 10'
combined_cols_help = 'columns that contain combined rows, default: 0,1,2'
due_date_text_help = 'text of the column as due date in excel file, default: \
Re-Validation Due:'
strip_text_help = 'text needs striped to get workcell name, can be multiple, \
default: equipment validation list'
high_user_help = 'high level emails list, can be multiple,default:Alex_Cheng@jabil.com'
dev_user_help = 'emails to make sure email sending out successfully, \
can be multiple,default:Alex_Cheng@jabil.com'
debug_help = 'only for development, send email to developer only'


@click.command(context_settings=dict(max_content_width=100))
@click.option('--folder', '-d', default='H:\\TE\\Proyekto\\configs',
              type=click.Path(exists=True), help=folder_help)
@click.option('--level', '-l', default=3, help=level_help)
@click.option('--level-maps', default='1:180,2:90,3:30',
              help=level_maps_help)
@click.option('--template', '-t', type=click.Path(exists=True),default='templates\one_month.html',
              help=template_help)
@click.option('--host', '-h', default='shateweb', help=host_help)
@click.option('--sender', '-s',
              default='Validation Expiring Alert$validation_alert@jabil.com',
              help=sender_help)
@click.option('--date-format', default='%d-%b-%Y', help=date_format_help)
@click.option('--date-columns', default='8,9',
              help=date_columns_help)
@click.option('--comment-columns', default='10', help=comment_columns_help)
@click.option('--combined-columns', default='0,1,2', help=combined_cols_help)
@click.option('--due-date-text', default='Re-Validation Due:',
              help=due_date_text_help)
@click.option('--strip-text', default=['equipment validation list'],
              multiple=True, help=strip_text_help)
@click.option('--high-user', multiple=True,default='Alex_Cheng@jabil.com', help=high_user_help)
@click.option('--dev-user', multiple=True,default='Alex_Cheng@jabil.com', help=dev_user_help)
@click.option('--debug/--no-debug', default=False, help=debug_help)
def run(folder, level, level_maps, template, host, sender, date_format,
        date_columns,comment_columns, combined_columns, due_date_text, strip_text, high_user,
        dev_user, debug):
    if sender:
        name, email = sender.split('$')
        sender = ((name, email))
    else:
        sender = (('Validation Expiring Alert', 'validation_alert@jabil.com'))

    level_days = {}
    for item in level_maps.split(','):
        l, days = item.split(':')
        level_days[int(l.strip())] = int(days.strip())
    
    dates_idx = [int(i.strip()) for i in date_columns.split(',')]
    comment_columns = [int(i.strip()) for i in comment_columns.split(',')]
    combined_columns = [int(i.strip()) for i in combined_columns.split(',')]
    due_date_text = due_date_text.strip().upper()
    strip_texts = [t.strip().upper() for t in strip_text]
    top_users = [t.strip() for t in high_user]
    dev_users = [t.strip() for t in dev_user]
    today = datetime.today()
    today = datetime(today.year, today.month, today.day)
    args = (today, int(level), level_days, due_date_text, date_format,
            dates_idx,comment_columns, combined_columns, debug)

    with open(template) as f:
        html_raw = T(f.read())

    subject_raw = T('Validation Expiring Alert - [{{ workcell }}]')
    if os.path.isfile(folder):
        files = [Path(folder)]
    else:
        d = Path(folder)
        files = d.files('*.xls') + d.files('*.xlsx')
    

    for file_path in files:
        if file_path.namebase.startswith('~'):
            continue

        if debug:
            click.echo('checking [{}]...'.format(file_path.name))

        rows, headers, to_list, comment, com_headers = parse_excel(file_path, *args)
        if rows is False:
            continue

        if not rows:
            continue

        if comment is False:
            continue
			
        if not comment:
		    continue
			
        wc = file_path.namebase.strip().upper()
        for st in strip_texts:
            wc = wc.split(st, 1)[0].strip()
	
        cc_list = []
        bcc = []
        if debug:
            to_list = dev_users
            cc_list = []
            bcc = []
        else:
            if top_users:
                cc_list = top_users

            bcc = dev_users
		
        msg = emails.html(subject=subject_raw,
                          html=html_raw,
                          mail_from=sender,
                          cc=cc_list,
                          bcc=bcc)
        msg.send(to=to_list, smtp=dict(host=host),
                 render=dict(file_path=file_path, workcell=wc, headers=headers,
                             fixtures=rows))
        msg.send(to=to_list, smtp=dict(host=host),
                render=dict(file_path=file_path, workcell=wc, headers=com_headers,
                            fixtures=comment))


def get_days(next_date, today, level, level_days):
    days = (next_date - today).days +365*3
    up_level = level + 1
    if up_level in level_days:
        if level_days[up_level] < days <= level_days[level]:
            return days
	
    else:
        if 0<days <= level_days[level]:
            return days

    return None


def convert_to_date(value):
    try:
        if isinstance(value, (float, int)):
            return datetime(*xldate_as_tuple(value, 0))
        elif isinstance(value, basestring):
            return parse(value.strip())

        return None
    except:
        return None


def get_index(values, keyword):
    for i, head in enumerate(values):
        if isinstance(head, basestring):
            if head.strip().upper().startswith(keyword):
                return i

    return None


def parse_excel(file_path, today, level, level_days, keyword, date_format,
                dates_idx=[], comment_idx=[],combined_columns=[], debug=False):
    rows = []
    headers = None
    to_list = []
    comment= []
    com_headers = None
    try:
        wb = open_workbook(file_path)
        ws = wb.sheet_by_index(0)
        idx = None
        all_rows = []
        for row in xrange(ws.nrows):
            values = ws.row_values(row)
            all_rows.append(values)
            if idx is None:
                i = get_index(values, keyword)
                if i is not None:
                    headers = values
                    idx = i

                continue

            if len(values) <= idx:
                continue

            try:
                if values[idx]:
                    next_date = convert_to_date(values[idx])
                    if not next_date:
                        continue

                else:
                    continue

                values[idx] = next_date.strftime(date_format)
                for idx_ in dates_idx:
                    if values[idx_]:
                        v_date = convert_to_date(values[idx_])
                        if v_date:
                            values[idx_] = v_date.strftime(date_format)

                for idx_ in combined_columns:
                    if not values[idx_]:
                        try:
                            values[idx_] = all_rows[-2][idx_]
                        except:
                            pass

                days = get_days(next_date, today, level, level_days)
                if days is not None:
                    rows.append(values)
                    # update updated row
                    all_rows[-1] = values
            except:
                if debug:
                    traceback.print_exc()

                continue
			
        ws = wb.sheet_by_index(2)
        #for row in xrange(ws.nrows):
            #values = ws.row_values(row)
            #if values:
				#comment.append(values)		
        idx = None
        all_rows = []
        for row in xrange(ws.nrows):
            values = ws.row_values(row)
            all_rows.append(values)
            if idx is None:
                i = get_index(values, keyword)
                if i is not None:
                    com_headers = values
                    idx = i

                continue

            if len(values) <= idx:
                continue

            try:
                if values[idx]:
                    next_date = convert_to_date(values[idx])
                    if not next_date:
                        continue

                else:
                    continue

                values[idx] = next_date.strftime(date_format)
                for idx_ in dates_idx:
                    if values[idx_]:
                        v_date = convert_to_date(values[idx_])
                        if v_date:
                            values[idx_] = v_date.strftime(date_format)

                for idx_ in combined_columns:
                    if not values[idx_]:
                        try:
                            values[idx_] = all_rows[-2][idx_]
                        except:
                            pass

                days = get_days(next_date, today, level, level_days)
                #if days is not None:
                comment.append(values)
                # update updated row
                all_rows[-1] = values
            except:
                if debug:
                    traceback.print_exc()

                continue
		
        ws = wb.sheet_by_index(1)
        for row in xrange(ws.nrows):
            values = [v for v in ws.row_values(row) if v]
            if values:
                for v in values:
                    email = ('{}'.format(v)).strip().replace('\xa0', '')
                    if email and '@' in email:
                        to_list.append(email)	
				
    except:
        if debug:
            traceback.print_exc()

        return False, None,  'Reading excel error: {}'.format(file_path)
	
    return rows, headers, to_list, comment, com_headers


if __name__ == '__main__':
    run()
