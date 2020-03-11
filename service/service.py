import json
import os
import logging;

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.client_request import ClientRequest
from office365.runtime.utilities.http_method import HttpMethod
from office365.runtime.utilities.request_options import RequestOptions
from office365.sharepoint.client_context import ClientContext

from flask import Flask, request, Response, abort

# Url for sharepoint site we woant to work on
URL = os.environ.get('SP_URL')

USERNAME = os.environ.get('SP_USERNAME')
PASSWORD = os.environ.get('SP_PASSWORD')

# Key for entity attribute containing name of list we want to work on
LIST_NAME = os.environ.get('SP_LIST_NAME', 'ListName')
# Key for entity attribute containing name of list item
# can be obtained by  sending GET to
# https://<tenant>.sharepoint.com/sites/<site>/_api/web/lists/GetByTitle('<list name>')/ListItemEntityTypeFullName
LIST_ITEM_NAME = os.environ.get('SP_LIST_ITEM_NAME', 'ListItemEntityTypeFullName')
LIST_SIZE = int(os.environ.get('SP_LIST_SIZE', '100'))

PORT = int(os.environ.get('PORT', '5000'))
LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO')

THREADS = int(os.environ.get('THREADS', '10'))

PROCESS_DELETED = os.environ.get('PROCESS_DELETED_ENTITIES', 'true').lower() == 'true'

APP = Flask(__name__)

if not URL or not USERNAME or not PASSWORD:
    logging.error("URL, USERNAME or PASSWORD not found.")
    exit(1)


@APP.route('/send-to-list', methods=['POST'])
def send_to_list():
    """
    Send list of entities to one or more sharepoint lists.
    Every entity must have:

    ListName property - which list it need to be sent

    ListItemEntityTypeFullName property - name of list item as it defined in Sharepoint
    one way to find it out is to call
    https://<name>.sharepoint.com/sites/<site name>/_api/lists/GetByTitle('<list name>')
    end find value of ListItemEntityTypeFullName property

    Keys property -list of entity keys which will be sent to SharePoint list

    Entity shall not have "status" attribute as it is used to storing result of operation
    If this attribute exist its value will be replaced!

    :return: list of processed entities where status attribute will be populated with result of operation
    """
    request_entities = request.get_json()

    def post_entities(entities: list):

        ctx_auth = AuthenticationContext(URL)

        ctx_auth.acquire_token_for_user(USERNAME, PASSWORD)
        if ctx_auth.provider.token:
            ctx = ClientContext(URL, ctx_auth)
        else:
            error = ctx_auth.get_last_error()
            logging.error(error)
            raise Exception(error)

        for _, entity in enumerate(entities):
            if entity['_deleted'] and not PROCESS_DELETED:
                logging.debug(f"entity {entity['_id']} marked as deleted and will not be processed")
                continue

            list_object = ctx.web.lists.get_by_title(entity[LIST_NAME])

            try:
                list_item_name = entity.get(LIST_ITEM_NAME)
                if list_item_name is None:
                    item_properties_metadata = {}
                else:
                    item_properties_metadata = {'__metadata': {'type': list_item_name}}
                keys_to_send = entity['Keys']
                values_to_send = {key: str(entity[key]) for key in keys_to_send}
                item_properties = {**item_properties_metadata, **values_to_send}

                existing_item = None
                if entity.get('ID'):
                    try:
                        existing_item = list_object.get_item_by_id(entity.get('ID'))
                        ctx.load(existing_item)
                        ctx.execute_query()
                    except Exception as ie:
                        logging.warning("Item lookup by ID resulted in an exception from Office 365 {}".format(ie))
                        if (hasattr(ie, 'code') and ie.code == "-2147024809, System.ArgumentException") or (
                                hasattr(ie,
                                        'message') and ie.message == "Item does not exist. It may have been deleted by another user."):
                            existing_item = None
                        else:
                            raise Exception from ie

                if not existing_item:
                    logging.info("Creating new item")
                    list_object.add_item(item_properties)
                    ctx.execute_query()
                else:
                    logging.info("Existing item found")
                    if entity.get('SHOULD_DELETE') is not None and bool(entity.get('SHOULD_DELETE')):
                        response = delete_list_item(ctx, entity[LIST_NAME], entity.get('ID'))
                    else:
                        response = update_list_item(ctx, entity[LIST_NAME], entity.get('ID'), values_to_send)
                    response.raise_for_status()

            except Exception as e:
                error_message = f"An exception occurred during processing of an entity: {e} ({json.dumps(entity)}"
                logging.error(error_message)
                raise Exception(error_message) from e

    post_entities(request_entities)
    return Response(status=200, response="{'status': 'success'}", mimetype='application/json')


@APP.route('/get-from-list/<list_name>', methods=['GET'])
def get_from_list(list_name):
    """
    Fetch list of entities from given sharepoint list
    :param list_name:
    :return:
    """

    def generate(entities):
        yield "["
        for index, entity in enumerate(entities):
            if index > 0:
                yield ","
            yield json.dumps(entity.properties)
        yield ']'

    ctx_auth = AuthenticationContext(URL)
    if ctx_auth.acquire_token_for_user(USERNAME, PASSWORD):
        ctx = ClientContext(URL, ctx_auth)
        list_object = ctx.web.lists.get_by_title(list_name)
        items = list_object.get_items().top(LIST_SIZE)
        ctx.load(items)
        ctx.execute_query()
        return Response(generate(items), mimetype='application/json')
    else:
        abort(500)


@APP.route('/get-site-users', methods=['GET'])
def get_site_users():
    """
    Fetch SharepointUsers users
    :return:
    """

    def generate(entities):
        yield "["
        for index, entity in enumerate(entities):
            if index > 0:
                yield ","
            yield json.dumps(entity.properties)
        yield ']'

    ctx_auth = AuthenticationContext(URL)
    if ctx_auth.acquire_token_for_user(USERNAME, PASSWORD):
        ctx = ClientContext(URL, ctx_auth)
        user_col = ctx.web.site_users
        ctx.load(user_col)
        ctx.execute_query()
    return Response(generate(user_col), mimetype='application/json')


def update_list_item(context, list_title, item_id, values_to_send):
    """
    Updates item with given id in given list with given properties
    :param context: auth context
    :param list_title: name of list
    :param item_id:
    :param values_to_send: dict with key-value pairs
    :return: requests/result object
    """
    request = ClientRequest(context)
    options = RequestOptions(
        "{2}/_api/web/lists/getbyTitle('{0}')/items({1})".format(list_title, item_id, URL))
    options.set_header('Accept', 'application/json; odata=nometadata')
    options.set_header('IF-MATCH', '*')
    options.set_header('X-HTTP-Method', 'MERGE')
    options.data = values_to_send
    options.method = HttpMethod.Post
    result = request.execute_request_direct(options)
    return result


def delete_list_item(context, list_title, item_id):
    """
    Deletes item with given id in given list
    :param context: auth context
    :param list_title: name of list
    :param item_id:
    :return: requests/response object
    """
    req = ClientRequest(context)
    options = RequestOptions(f"{URL}/_api/web/lists/getbyTitle('{list_title}')/items({item_id})")
    options.set_header('Accept', 'application/json; odata=nometadata')
    options.set_header('IF-MATCH', '*')
    options.set_header('X-HTTP-Method', 'DELETE')
    options.method = HttpMethod.Post
    res = req.execute_request_direct(options)
    return res


if __name__ == '__main__':
    logging.basicConfig(level=logging.getLevelName(LOG_LEVEL))

    IS_DEBUG_ENABLED = logging.getLogger().isEnabledFor(logging.DEBUG)

    if IS_DEBUG_ENABLED:
        APP.run(debug=IS_DEBUG_ENABLED, host='0.0.0.0', port=PORT)
    else:
        import cherrypy

        cherrypy.tree.graft(APP, '/')
        cherrypy.config.update({
            'environment': 'production',
            'engine.autoreload_on': True,
            'log.screen': False,
            'server.socket_port': PORT,
            'server.socket_host': '0.0.0.0',
            'server.thread_pool': THREADS,
            'server.max_request_body_size': 0
        })

        cherrypy.engine.start()
        cherrypy.engine.block()
