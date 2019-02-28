#!/usr/bin/python3

import tornado.ioloop
import tornado.netutil
import tornado.httpclient

from urllib.parse import urlencode

import json


class UriTools:
    def __init__(self, service_uri):
        self._service_uri = service_uri

    async def parse(self, uri, decorator=None):
        http_client = tornado.httpclient.AsyncHTTPClient()
        try:

            params = {
                'uri': uri,
                'decorator': decorator
            }
            call_uri = "{0}/clean?{1}".format(self._service_uri, urlencode(params))

            response = await http_client.fetch(call_uri)

            return response.body.decode("utf-8")

        except tornado.httpclient.HTTPError as e:
            if e.code == 404:
                return None
            raise e
        except Exception as e:
            raise e

    async def parse_as_json(self, uri):
        res = await self.parse(uri, decorator="json")

        return json.loads(res)


if __name__ == "__main__":
    print("This is a module.")
