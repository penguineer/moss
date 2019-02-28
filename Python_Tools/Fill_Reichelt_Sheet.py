#!/usr/bin/python3

import argparse

from uritools import UriTools

import tornado.ioloop
import tornado.netutil
import tornado.httpclient

from urllib.parse import urlencode

from pyquery import PyQuery

import pyoo

import locale


def tag_to_discounts(tag):
    discounts = dict()

    i = 0
    while True:
        li = tag('li').eq(i)
        if not li:
            break
        i += 1

        s_pr = li('span').eq(0).text()
        price = s_pr[0:-2]
        amount = li('span').eq(1).text()
        discounts[int(amount)] = price

    return discounts


async def reichelt_search(artid):
    http_client = tornado.httpclient.AsyncHTTPClient()
    try:
        params = {
            'SEARCH': artid
        }
        body = urlencode(params)
        uri = "https://www.reichelt.de/index.html?ACTION=446&LA=0"

        response = await http_client.fetch(uri, method="POST", body=body)

        link = None
        price = None
        discounts = None

        pq = PyQuery(response.body)
        i = 0
        while True:
            tag = pq('div.al_gallery_article').eq(i)
            if not tag:
                break
            i += 1

            meta = tag('meta[itemprop="productID"]')
            if artid == meta.attr("content"):
                link = tag('a.al_artinfo_link').attr("href")
                price = tag('span[itemprop="price"]').text()

                dsc = tag('ul.discounts')
                if dsc:
                    discounts = tag_to_discounts(dsc)

                break

        return link, price, discounts
    except Exception as e:
        raise e


async def open_reichelt_document(path, host="localhost", port=2002):
    desktop = pyoo.Desktop(host, int(port))
    doc = desktop.open_spreadsheet(path)

    return doc


async def reichelt_sheet_to_csv(document):
    sheet = document.sheets[0]

    csv = ""

    i = 1
    while sheet[i][0].value:
        artid = sheet[i][0].value
        amount = int(sheet[i][2].value)
        csv += "{art};{amount}\n".format(art=artid, amount=amount)
        i += 1

    return csv


def next_highest(seq, x):
    return min([(i - x, i) for i in seq if x <= i] or [(0, None)])[1]


def next_lowest(seq, x):
    return min([(x-i, i) for i in seq if x >= i] or [(0, None)])[1]


def get_discounted_price(price, amount, discounts):
    if not amount:
        raise ValueError("Amount must be provided!")

    if discounts:
        d = [*discounts]
        d.sort()
        am = next_lowest(d, amount)
        price = discounts[am]

    return price


async def reichelt_sheet_complete(document, uritools):
    sheet = document.sheets[0]

    locale.setlocale(locale.LC_NUMERIC, 'de_DE.UTF-8')

    i = 1
    while sheet[i][0].value:
        artid_cell = sheet[i][0]
        desc_cell = sheet[i][1]
        amount_cell = sheet[i][2]
        price_cell = sheet[i][3]
        if not desc_cell.value or not price_cell.value:
            link, price, discounts = await reichelt_search(artid_cell.value)
            price = get_discounted_price(price, int(amount_cell.value), discounts)

            if link:
                res = await uritools.parse_as_json(link)
                res['price'] = locale.atof(price)

                desc_cell.value = res['name']
                # check price
                if not price_cell.value:
                    price_cell.value = res['price']
            else:
                artid_cell.text_color = 0xFF0000
        i += 1


async def _main():
    parser = argparse.ArgumentParser(
        description="Complete Reichelt mass order sheet and export basket CSV")
    parser.add_argument("--service", help="The uritools service URI")
    parser.add_argument("--desktop_host", help="LibreOffice desktop host (localhost)", default="localhost")
    parser.add_argument("--desktop_port", help="LibreOffice desktop port (2002)", default=2002)
    parser.add_argument("--sheet", help="Path to LibreOffice sheet")
    args = parser.parse_args()

    service = args.service

    uritools = UriTools(service)

    doc = await open_reichelt_document(args.sheet, host=args.desktop_host, port=int(args.desktop_port))
    csv = await reichelt_sheet_to_csv(doc)
    print(csv)

    await reichelt_sheet_complete(doc, uritools)

    # Don't close so that the user may save the document
#    doc.close()

if __name__ == "__main__":
    tornado.ioloop.IOLoop.current().run_sync(_main)
