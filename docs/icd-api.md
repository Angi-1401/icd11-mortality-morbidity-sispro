**Disclaimer: All information presented in this document comes directly from the official [ICD API Documentation](https://icd.who.int/docs/icd-api/APIDoc-Version2/).**

# ICD API

ICD API allows programmatic access to the International Classification of Diseases (ICD). It is an HTTP based REST API. You may use [this site](https://icd.who.int/icdapi) to access up to date documentation on using the API as well as managing the keys needed for using the API.

All communication made to the APIs are encrypted (i.e. only https is allowed). All http requests will be automatically redirected to the https variants.

Even though there is this automated redirection, we recommend directly calling to the https endpoints as this will work faster

## API Access

In order to be able to use the ICD APIs, first you need to create an account on the ICD API Home page: https://icd.who.int/icdapi

The APIs use OAuth 2 client credentials for authentication. Once you register and login to this site, you could get a client id and client secret to be used for authentication. This is done by clicking on the View API access key link.

Token Endpoint for the service is located at: 

```url
https://icdaccessmanagement.who.int/connect/token
```

More information on authentication can be found in the [ICD API Authentication](https://icd.who.int/docs/icd-api/API-Authentication/) document

All communication made to the access management site and the APIs are encrypted (i.e. only https is allowed) However if you refer to the http variants of the URLs, they will be automatically redirected.

## How to obtain an SECRET_ID and SECRET_KEY from the ICD API

1. Access the ICD API Home page: https://icd.who.int/icdapi
2. Create an account and login to the site.
3. Click on the View API access key link.
4. Retrieve your credentials and store them in a secure place.