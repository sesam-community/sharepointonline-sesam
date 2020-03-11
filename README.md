# sharepointonline-sesam
Sharepoint online source+sink for Sesam.io

[![Build Status](https://travis-ci.org/sesam-community/sharepointonline-sesam.svg?branch=master)](https://travis-ci.org/sesam-community/sharepointonline-sesam)

system setup
```json
{
  "_id": "sharepoint-sink",
  "type": "system:microservice",
  "docker": {
    "environment": {
      "SP_PASSWORD": "$SECRET(sp-password)",
      "SP_URL": "$ENV(sp-url)",
      "SP_USERNAME": "$ENV(sp-username)",
      "SP_LIST_SIZE": "100"
    },
    "image": "<docker image name>",
    "port": 5000
  }
}

```

pipe setup
```json
{
  "_id": "<pipe id>",
  "type": "pipe",
  "source": {
    "type": "dataset",
    "dataset": "<source dataset>"
  },
  "transform": [{
    "type": "dtl",
    "rules": {
      "default": [
        ["add", "::Key1", "_S.<source property key>"],
        ["add", "::Key2", "_S.<sourceproperty key>"],
        ["add", "::ListName", "<SharePointListName>"],
        ["add", "::ListItemEntityTypeFullName", "SP.Data.<SharePoint List Item>"],
        ["add", "::Keys",
          ["list", "Key1", "Key2"] <-- which attributes are we gonna send (we don't need to send all, SP will return error on unrecognized fields)
        ],
        ["add", "::SHOULD_DELETE", true] <-- if want to delete an entity instead of updating it
      ]
    }
  },
    {
    "type": "http",
    "system": "sharepoint-sink",
    "url": "/send-to-list"
  }],
  "pump": {
    "schedule_interval": 15
  }
}
```
#### environment variables to controll the service
* `SP_LIST_SIZE` is defaulted to 100 elements, but you can bump the value up as needed.
* `PORT` default 5000
* `LOG_LEVEL` default INFO
* `THREADS` CherryPy thread pool size defaulted to 10
* `PROCESS_DELETED` - if `_deleted` entities should be processed as all other or excluded from processing, true by default (to preserve compatability with projects where this service is used)
