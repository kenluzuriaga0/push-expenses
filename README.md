# push-expenses

Registra automáticamente los gastos de tarjeta de crédito en **Actual Budget** leyendo los emails de notificación de Produbanco en **Outlook/Hotmail** vía Microsoft Graph API.

---

## Cómo funciona

1. Autentica en Outlook vía OAuth2 (device code flow — login en navegador la primera vez)
2. Lee emails del inbox usando Microsoft Graph API
3. Filtra los que coinciden con el regex configurado (ej: _"Consumo Tarjeta de Crédito por USD 3.78"_)
4. Extrae el monto del asunto y el establecimiento del cuerpo del email
5. Asigna un payee automático si el establecimiento coincide con alguna regla
6. Verifica duplicados en Actual Budget antes de insertar (ventana de ±3 días)
7. Inserta la transacción en la cuenta configurada
8. Marca el email con una categoría en Outlook para no reprocesarlo en futuras ejecuciones

---

## Requisitos

- Node.js 18+
- Instancia de [Actual Budget](https://actualbudget.org/) corriendo (self-hosted)
- Cuenta de Outlook/Hotmail personal

---

## Instalación

```bash
npm install
cp .env.example .env
# Edita .env con tus valores
npm start
```

---

## Configurar Azure para acceder a Outlook

> Costo: **$0**. El registro de apps en Azure es gratuito para cuentas personales.

### 1. Registrar la app

1. Ve a [portal.azure.com](https://portal.azure.com) e inicia sesión con tu cuenta Microsoft personal
2. Busca **"App registrations"** → **New registration**
3. Completa el formulario:
   - **Name:** `push-expenses` (o cualquier nombre)
   - **Supported account types:** _Personal Microsoft accounts only_
   - **Redirect URI:** dejar en blanco
4. Click **Register**
5. En la pantalla **Overview**, copia el **Application (client) ID** → va a `AZURE_CLIENT_ID` en tu `.env`

### 2. Habilitar Device Code Flow

Sin este paso, la autenticación falla con `invalid_grant`.

1. En tu app → **Authentication**
2. Baja hasta **Advanced settings**
3. **Allow public client flows** → cambia a **Yes**
4. **Save**

### 3. Agregar permisos de email

1. En tu app → **API permissions** → **Add a permission**
2. Selecciona **Microsoft Graph** → **Delegated permissions**
3. Busca y agrega:
   - `Mail.Read` — para leer el inbox
   - `Mail.ReadWrite` — para marcar emails como procesados
4. Click **Add permissions**

> Para cuentas personales no necesitas "Grant admin consent". El consentimiento se otorga automáticamente la primera vez que haces login.

---

## Primer login

La primera vez que corres `npm start`, el script imprime:

```
Abre: https://microsoft.com/devicelogin
Código: ABC123XY
```

Abre el link, ingresa el código y haz login con tu cuenta Hotmail. El token se guarda en `.token-cache.json` y se renueva automáticamente. Las siguientes ejecuciones son completamente automáticas.

---

## Obtener el Sync ID de Actual Budget

1. Abre tu Actual Budget en el navegador
2. Ve a **Settings** → **Show advanced settings**
3. Copia el valor de **Sync ID** → va a `ACTUAL_BUDGET_ID` en tu `.env`

---

## Variables de entorno

| Variable | Requerida | Descripción |
|---|---|---|
| `ACTUAL_SERVER_URL` | ✓ | URL de tu servidor Actual Budget (ej: `http://192.168.1.10:5006`) |
| `ACTUAL_PASSWORD` | ✓ | Contraseña del servidor |
| `ACTUAL_BUDGET_ID` | ✓ | Sync ID del presupuesto |
| `ACTUAL_ACCOUNT_NAME` | ✓ | Nombre exacto de la cuenta destino en Actual |
| `AZURE_CLIENT_ID` | ✓ | Application (client) ID de tu app en Azure |
| `ACTUAL_DATA_DIR` | — | Directorio de caché local (default: `.actual-data`) |
| `EMAIL_SUBJECT_REGEX` | — | Regex para filtrar emails; el último grupo numérico es el monto |
| `EMAIL_FROM_DATE` | — | Buscar emails desde esta fecha `YYYY-MM-DD`; ignora `EMAIL_LIMIT` |
| `EMAIL_LIMIT` | — | Últimos N emails a leer si no se define `EMAIL_FROM_DATE` (default: `50`) |
| `PROCESSED_LABEL` | — | Categoría que se asigna al email procesado (default: `ActualBudget`) |

---

## Personalizar el regex del asunto

El valor por defecto captura asuntos como _"Consumo Tarjeta de Crédito por USD 3.78"_:

```
EMAIL_SUBJECT_REGEX=Consumo Tarjeta de Crédito por (?:USD|\$) ([\d.]+)
```

El **último grupo numérico** capturado se usa como monto. Ejemplos para otros formatos:

```
# "Compra aprobada por $25.00"
EMAIL_SUBJECT_REGEX=Compra aprobada por \$([\d.]+)

# "Transaccion TC: USD 100.50"
EMAIL_SUBJECT_REGEX=Transaccion TC: USD ([\d.]+)
```

---

## Reglas de payee automático

En `src/outlook.ts` hay un array `PAYEE_RULES` que mapea patrones del nombre del establecimiento a payees de Actual Budget. Actual Budget aplica sus propias reglas de categoría según el payee, así que basta con asignar el payee correcto.

```typescript
const PAYEE_RULES = [
  { pattern: /UBER/i,        payee: 'Uber' },
  { pattern: /COMISARIATO/i, payee: 'Comisariato' },
];
```

Para agregar una regla nueva, añade una línea al array siguiendo el mismo formato.

---

## Comandos

```bash
npm start       # procesar emails nuevos e insertar en Actual Budget
npm run list    # listar las últimas 10 transacciones del servidor (para verificar)
npm run build   # compilar TypeScript a dist/
```

---

## Estructura del proyecto

```
src/
├── index.ts    — configuración y orquestación principal
├── outlook.ts  — auth OAuth2, Graph API, parsing de emails, payee rules
├── actual.ts   — conexión a Actual Budget, inserción, detección de duplicados
└── list.ts     — utilidad para listar transacciones y verificar el sync
```

---

## Por qué no se usa IMAP

La opción más simple sería conectarse vía IMAP con usuario y contraseña. **No funciona.** Microsoft deprecó la autenticación básica (usuario + contraseña / app passwords) para cuentas personales de Outlook.com en septiembre de 2025. El servidor IMAP responde con `AUTHENTICATE failed` aunque tengas IMAP habilitado y hayas generado una contraseña de aplicación.

La única forma de acceder al correo de una cuenta personal de Hotmail/Outlook.com es mediante **OAuth2 con Microsoft Graph API**, que es lo que hace este proyecto.
