#!/usr/bin/env python3

import site
import sys

new_sys_path = set(sys.path)
for path in site.getsitepackages():
    if path in new_sys_path:
        new_sys_path.remove(path)
sys.path[0:] = new_sys_path

import os.path

if __name__ == '__main__':
    # sys.path.append('/usr/lib/libreoffice/program')
    sys.path.append('/usr/lib/python3/dist-packages')
    import uno
    import unohelper
    sys.path.remove('/usr/lib/python3/dist-packages')
    sys.path.append(os.path.join(os.path.dirname(__file__),
                                 'pythonpath'))


import argparse
import collections
import csv
import io
import logging
import os
import traceback
import urllib.request

import boto3
import uno       # NOQA
import unohelper # NOQA
from com.sun.star.sheet.CellFlags import VALUE
from com.sun.star.sheet.CellFlags import DATETIME
from com.sun.star.sheet.CellFlags import STRING
from com.sun.star.sheet.CellFlags import ANNOTATION
from com.sun.star.sheet.CellFlags import FORMULA
from com.sun.star.sheet.CellFlags import HARDATTR
from com.sun.star.sheet.CellFlags import STYLES
from com.sun.star.sheet.CellFlags import OBJECTS
from com.sun.star.sheet.CellFlags import EDITATTR
from com.sun.star.sheet.CellFlags import FORMATTED
from com.sun.star.task import XJob
from com.sun.star.container import NoSuchElementException
from com.sun.star.awt.FontWeight import BOLD

BOOTSTRAP_REGION = 'us-east-1'

ALL_CELL_FLAGS = (VALUE | DATETIME | STRING | ANNOTATION | FORMULA |
                  HARDATTR | STYLES | OBJECTS | EDITATTR | FORMATTED)

g_ImplementationHelper = unohelper.ImplementationHelper()


def _get_or_create_sheet(doc, sheet_name):
    try:
        sheet = doc.Sheets.getByName(sheet_name)
    except NoSuchElementException:
        doc.Sheets.insertNewByName(sheet_name, doc.Sheets.Count)
        sheet = doc.Sheets.getByName(sheet_name)
    return sheet


class ImportEC2Pricing(unohelper.Base, XJob):
    EC2_PRICE_LIST_URL = ('https://pricing.us-east-1.amazonaws.com/offers/'
                          'v1.0/aws/AmazonEC2/current/index.csv')
    PRICE_LIST_HEADER_FIELDS = [
        'FormatVersion', 'Disclaimer', 'Publication Date', 'Version',
        'OfferCode']
    PRICING_SHEET_NAME = 'pricing data'
    PRICING_METADATA_SHEET_NAME = 'pricing metadata'
    NUM_DATA_ROWS_TO_LOAD = 1000

    def __init__(self, ctx):
        self.ctx = ctx

    def _get_pricing_metadata(self, pricing_csv_text):
        metadata = collections.OrderedDict()
        buf = io.StringIO()
        for i in range(len(self.PRICE_LIST_HEADER_FIELDS)):
            current_char = None
            while current_char != '\n':
                current_char = pricing_csv_text.read(1)
                buf.write(current_char)

        buf.seek(0)
        csv_data = csv.reader(buf)
        for row, expected_header_field in zip(csv_data,
                                              self.PRICE_LIST_HEADER_FIELDS):
            if row[0] != expected_header_field:
                raise Exception('Incorrect header field detected: {}'.format(
                    row[0]))

            metadata[row[0]] = row[1]

        return metadata

    def _get_pricing_lines(self, pricing_lines_csv_text):
        data = []
        csv_data = csv.reader(pricing_lines_csv_text)

        csv_iter = iter(csv_data)
        header_row = next(csv_iter)
        for row in csv_iter:
            data.append(tuple(row))

        return tuple(header_row), tuple(data)

    def _get_pricing_data(self):
        with urllib.request.urlopen(self.EC2_PRICE_LIST_URL) as r, \
                io.TextIOWrapper(
                    r,
                    encoding=r.headers.get_content_charset('utf-8')) \
                as pricing_text:
            metadata = self._get_pricing_metadata(pricing_text)
            headers, data = self._get_pricing_lines(pricing_text)

        return metadata, headers, data

    def _update_pricing_metadata_sheet(self, sheet, metadata):
        sheet.clearContents(ALL_CELL_FLAGS)
        for row, row_key in enumerate(metadata):
            cell_range = sheet.getCellRangeByPosition(0, row, 1, row)
            cell_range.setDataArray(((row_key, metadata[row_key]),))

    def _update_pricing_data_sheet(
            self, sheet, headers, data_rows, status_indicator):
        sheet.clearContents(ALL_CELL_FLAGS)
        cell_range = sheet.getCellRangeByPosition(0, 0, len(headers) - 1, 0)
        cell_range.setDataArray((tuple(headers),))
        cell_range.CharWeight = BOLD

        all_data_range = sheet.getCellRangeByPosition(
            0, 1, len(headers)-1, len(data_rows))
        print('len(data_rows): {}'.format(len(data_rows)))
        status_indicator.start('Loading pricing data...', len(data_rows))
        for start_row in range(0, len(data_rows) - self.NUM_DATA_ROWS_TO_LOAD,
                               self.NUM_DATA_ROWS_TO_LOAD):
            print('loading start row: {}'.format(start_row))
            end_row = start_row + self.NUM_DATA_ROWS_TO_LOAD
            current_range = all_data_range.getCellRangeByPosition(
                0, start_row, len(headers)-1, end_row - 1)
            current_range.setDataArray(data_rows[start_row:end_row])
            status_indicator.setValue(end_row)
        else:
            start_row = start_row + self.NUM_DATA_ROWS_TO_LOAD
            if start_row < len(data_rows):
                print('loading start row: {}'.format(start_row))
                end_row = len(data_rows)
                cell_range = all_data_range.getCellRangeByPosition(
                    0, start_row, len(headers)-1, end_row - 1)
                cell_range.setDataArray(data_rows[start_row:end_row])
                status_indicator.setValue(end_row)
        status_indicator.end()
        print('Row loaded: {}'.format(end_row))

    def execute(self, args):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)

        doc = desktop.getCurrentComponent()

        pricing_data_sheet = _get_or_create_sheet(doc, self.PRICING_SHEET_NAME)
        pricing_metadata_sheet = _get_or_create_sheet(
            doc, self.PRICING_METADATA_SHEET_NAME)

        pricing_metadata, pricing_data_headers, pricing_data = (
            self._get_pricing_data())

        try:
            self._update_pricing_metadata_sheet(
                pricing_metadata_sheet, pricing_metadata)
            status_indicator = (
                doc.getCurrentController().getFrame().createStatusIndicator())
            import datetime
            start = datetime.datetime.now()
            self._update_pricing_data_sheet(
                pricing_data_sheet, pricing_data_headers, pricing_data,
                status_indicator)
            end = datetime.datetime.now()
            print(end-start)
        except:
            print(traceback.format_exc())


g_ImplementationHelper.addImplementation(
    ImportEC2Pricing,
    'org.penguintechs.wt.libreoffice.ImportEC2Pricing',
    ('com.sun.star.task.Job',),)


class ImportEC2InstanceData(unohelper.Base, XJob):
    INSTANCE_COUNTS_SHEET_NAME = 'current instance counts'
    INSTANCE_COUNTS_SHEET_HEADERS = ('az', 'service', 'instance_type', 'count')
    RESERVED_INSTANCE_COUNTS_SHEET_NAME = 'current reserved instance counts'
    RESERVED_INSTANCE_COUNTS_SHEET_HEADERS = ('az', 'instance_type', 'count')

    def __init__(self, ctx):
        self.ctx = ctx

    def _get_all_regions(self):
        print('Getting list of regions')
        client = boto3.client('ec2', region_name=BOOTSTRAP_REGION)
        regions_dict = client.describe_regions()['Regions']
        regions = [i['RegionName'] for i in regions_dict]
        print('Found regions: {}'.format(regions))
        return regions

    def _get_instance_type_counts_for_region(self, resource):
        counts = {}
        for instance in resource.instances.all():
            instance_type = instance.instance_type
            az = instance.placement['AvailabilityZone']
            service = 'unknown'
            if instance.tags is not None:
                for tag_dict in instance.tags:
                    if tag_dict['Key'] == 'service':
                        service = tag_dict['Value']
            counts[(az, service, instance_type)] = counts.get(
                (az, service, instance_type), 0) + 1
        return counts

    def _update_instance_counts_sheet(self, doc, instance_counts):
        sheet = _get_or_create_sheet(doc, self.INSTANCE_COUNTS_SHEET_NAME)
        sheet.clearContents(ALL_CELL_FLAGS)

        headers = self.INSTANCE_COUNTS_SHEET_HEADERS
        headers_range = sheet.getCellRangeByPosition(0, 0, len(headers) - 1, 0)
        headers_range.setDataArray((headers,))
        headers_range.CharWeight = BOLD

        data_range = sheet.getCellRangeByPosition(
            0, 1, len(headers) - 1, len(instance_counts))
        for row, ((az, service, instance_type), count) in enumerate(
                instance_counts.items()):
            row_range = data_range.getCellRangeByPosition(
                0, row, len(headers) - 1, row)
            row_range.setDataArray(((az, service, instance_type, count),))

    def _get_reserved_instance_counts_for_region(self, client):
        ris = client.describe_reserved_instances(
            Filters=[{'Name': 'state',
                      'Values': ['payment-pending', 'active']}])
        ri_counts = {}
        for ri in ris['ReservedInstances']:
            az = ri['AvailabilityZone']
            instance_type = ri['InstanceType']

            ri_counts[(az, instance_type)] = (
                ri_counts.get((az, instance_type), 0) + ri['InstanceCount'])
        return ri_counts

    def _update_reserved_instance_counts_sheet(self, doc, ri_counts):
        sheet = _get_or_create_sheet(
            doc, self.RESERVED_INSTANCE_COUNTS_SHEET_NAME)
        sheet.clearContents(ALL_CELL_FLAGS)

        headers = self.RESERVED_INSTANCE_COUNTS_SHEET_HEADERS
        headers_range = sheet.getCellRangeByPosition(0, 0, len(headers) - 1, 0)
        headers_range.setDataArray((headers,))
        headers_range.CharWeight = BOLD

        data_range = sheet.getCellRangeByPosition(
            0, 1, len(headers) - 1, len(ri_counts))
        for row, ((az, instance_type), count) in enumerate(ri_counts.items()):
            row_range = data_range.getCellRangeByPosition(
                0, row, len(headers) - 1, row)
            row_range.setDataArray(((az, instance_type, count),))

    def execute(self, args):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)

        doc = desktop.getCurrentComponent()

        instance_counts = {}
        ri_counts = {}
        for region in self._get_all_regions():
            print('Getting instance data from {}'.format(region))
            resource = boto3.resource('ec2', region_name=region)
            client = boto3.client('ec2', region_name=region)
            instance_counts.update(
                self._get_instance_type_counts_for_region(resource))
            ri_counts.update(
                self._get_reserved_instance_counts_for_region(client))

        self._update_instance_counts_sheet(doc, instance_counts)
        self._update_reserved_instance_counts_sheet(doc, ri_counts)


g_ImplementationHelper.addImplementation(
    ImportEC2InstanceData,
    'org.penguintechs.wt.libreoffice.ImportEC2InstanceData',
    ('com.sun.star.task.Job',),)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('command')
    return parser.parse_args()


def main():
    args = parse_args()

    connect_string = 'socket,host=localhost,port=2002;urp'

    # Start OpenOffice.org, listen for connections and open testing document
    os.system(
        "/usr/bin/libreoffice '--accept={};' --calc ./costing_test_doc.ods &"
        .format(connect_string))

    # Get local context info
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)

    ctx = None

    # Wait until the OO.o starts and connection is established
    while ctx is None:
        try:
            ctx = resolver.resolve(
                "uno:{};StarOffice.ComponentContext".format(connect_string))
        except:
            pass

    # Execute our job
    if args.command == 'import_pricing':
        blah = ImportEC2Pricing(ctx)
        blah.execute(())
    elif args.command == 'import_instance_data':
        blah2 = ImportEC2InstanceData(ctx)
        blah2.execute(())

if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    logging.getLogger('boto3').setLevel(logging.INFO)
    logging.getLogger('botocore').setLevel(logging.INFO)
    main()
