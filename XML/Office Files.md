#Office Files

Il nuovo formato dei documenti Office (docx, xlsx, pptx) consiste in una cartella zip che contiene dati strutturati in XML.

##Powerpoint

In un documento `pptx`:

* ppt
  * media (lì si trovano le immagini ed i filmati)  

##Proprietà di un documento Excel

Le proprietà di un documento sono conservate all'interno della cartella `/customXml`.
Lì si trovano tre coppie di documenti:
* `item1.xml` e `itemProps1.xml`
* `item2.xml` e `itemProps2.xml`
* `item3.xml` e `itemProps3.xml`


###item1
Nel documento `item1.xml` vengono riassunte tutte le proprietà

Esempio:
```
<xsd:element name="_x0034_" ma:index="10" nillable="true" ma:displayName="4" ma:internalName="_x0034_">
```

`name="_x0034_"`: il nome della proprietà. `_x003` indica che si tratta di un numero. Infatti...
`ma:displayName="4"`: il nome da mostrare è semplicemente "4".
`nillable="true"`: la proprietà non è indispensabile ed il campo può quindi essere lasciato vuoto. In caso contrario il predicato non è presente con valore `false`, è semplicemente assente.

####Tipo di valore

`<xsd:restriction base="dms:Text">`: testo; concluso da un `</xsd:restriction>`, nei casi visti seguito da `<xsd:maxLength value="255"/>`.
`<xsd:restriction base="dms:Number"/>`: numero (unico caso visto di attributo con auto chisuura).
`<xsd:restriction base="dms:Choice">`: scelta tra i valori proposti; concluso da un `</xsd:restriction>`.

###item2.xml

Esempio di documento
```
<?mso-contentType?><FormTemplates xmlns="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms"><Display>DocumentLibraryForm</Display><Edit>DocumentLibraryForm</Edit><New>DocumentLibraryForm</New></FormTemplates>
```

###item3.xml

Dichiarazione delle proprietà:

```
<p:properties xmlns:p="http://schemas.microsoft.com/office/2006/metadata/properties" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:pc="http://schemas.microsoft.com/office/infopath/2007/PartnerControls">
```

E dopo i valori delle proprietà, in ordine di inserimento(?).
```
<_x0034_ xmlns="49b6cb60-5abe-41d9-9a4c-b4633802e7c7">zzz</_x0034_
```

##Fonti
https://msdn.microsoft.com/en-us/library/dd208104.aspx
https://msdn.microsoft.com/fr-fr/library/aa338205(v=office.12).aspx
