https://github.com/google/clasp

https://hackernoon.com/writing-google-apps-script-code-locally-in-vscode

- npx clasp -P src/ pull
- npx clasp -P src/ push

# PASOS PARA AGREGAR APP SCRIPTS A PLANILLAS VIEJAS:

1 - convertir de excel a a planilla de google.

2 - ir a extensiones -> app script.

3 - click en el + de Bibliotecas.

4 - pegar id de la biblioteca:
`<KEY>`

5 - click en Buscar, y cuando encuentre click en Agregar/Aceptar.

6 - en la parte del codigo copiar el codigo del script:

```
function myFunction() {
  contaduriadevlights.updateResumen()
  contaduriadevlights.updateTotales()
}
```

7 - click en ejecutar.

#

<ID>

contaduriadevlights

function myFunction() {
contaduriadevlights.updateResumen()
contaduriadevlights.updateTotales()
}

#

1 - ir a extensiones -> app script.

2- click en el + de Bibliotecas.

3 - pegar id de la biblioteca:
`<KEY>`

5 - click en Buscar, y cuando encuentre click en version, click en 1, y click en Agregar/Aceptar.

6 - en la parte del codigo copiar el codigo del script:

```
function myFunction() {
  sistemadevlights.myFunction()
}
```

7 - click en ejecutar.
