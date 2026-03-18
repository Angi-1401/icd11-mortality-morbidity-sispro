**Aviso: Toda la información presentada en este documento proviene directamente de la [Documentación oficial de la API ICD](https://icd.who.int/docs/icd-api/APIDoc-Version2/).**

# API ICD

La API ICD permite el acceso programático a la Clasificación Internacional de Enfermedades (ICD). Es una API REST basada en HTTP. Puede usar [este sitio](https://icd.who.int/icdapi) para acceder a la documentación actualizada sobre el uso de la API, así como para gestionar las claves necesarias para su uso.

Toda la comunicación realizada con las APIs está cifrada (es decir, solo se permite https). Todas las solicitudes http se redirigirán automáticamente a sus variantes https.

Aunque existe esta redirección automática, recomendamos llamar directamente a los endpoints https ya que funcionará más rápido.

## Acceso a la API

Para poder usar las APIs de ICD, primero debe crear una cuenta en la página principal de la API ICD: https://icd.who.int/icdapi

Las APIs utilizan credenciales de cliente OAuth 2 para la autenticación. Una vez que se registre e inicie sesión en este sitio, podrá obtener un client id y client secret para autenticarse. Esto se realiza haciendo clic en el enlace "View API access key".

El endpoint de token para el servicio se encuentra en:

```url
https://icdaccessmanagement.who.int/connect/token
```

Puede encontrar más información sobre autenticación en el documento [ICD API Authentication](https://icd.who.int/docs/icd-api/API-Authentication/)

Toda la comunicación realizada con el sitio de gestión de acceso y las APIs está cifrada (solo https). Sin embargo, si utiliza las variantes http de las URLs, serán redirigidas automáticamente.

## Cómo obtener un SECRET_ID y SECRET_KEY de la API ICD

1. Acceda a la página principal de la API ICD: https://icd.who.int/icdapi
2. Cree una cuenta e inicie sesión en el sitio.
3. Haga clic en el enlace "View API access key".
4. Recupere sus credenciales y guárdelas en un lugar seguro.
