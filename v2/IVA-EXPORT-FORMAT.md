# The DRORIVA declaration file — reverse-engineered, not guessed

What `DRORIVA.app` opens and writes, so IVA rows can be exported instead of
retyped. **Nothing is built yet.** This is the format spec; the generator is a
separate piece of work.

Established 13 Aug 2026 from a one-row declaration saved by the app itself,
because the published XSDs are not sufficient — see the traps below. The sample
is **not** in this repo: it embeds a real invoice PDF and a real supplier NIF.

## What the app is

`/Applications/DRORIVA.app` — Portuguese Tax Authority, main class
`pt.at.simplex130.Simplex130Application`, Java 8 Swing packaged with install4j.
It submits to:

```
https://www.portaldasfinancas.gov.pt/pt/externalws/drorivaws/entregaDeclORIVAOffline.action
```

The published schema is on this machine at
`~/Downloads/Suporte_Informatico_Restituicoes_IVA_OR/` — `Simplex130.xsd`,
`types.xsd`, `catalogs.xsd`, all dated 2017.

## The format

```xml
<?xml version='1.0' encoding='ISO-8859-1'?>

<Simplex130 xmlns="http://www.dgci.gov.pt/2013/Simplex130" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.dgci.gov.pt/2013/Simplex130 Simplex130.xsd" versao="1" >
  <Rosto>
    <Quadro01>
      <NifEntidade >NNNNNNNNN</NifEntidade>
      <TipoEntidade >d</TipoEntidade>
      <NifFuncionario >NNNNNNNNN</NifFuncionario>
    </Quadro01>
    <Quadro02>
      <Quadro02T1 >
        <Quadro02T1-Linha numero="1">
          <NumeroFatura >FAC33484</NumeroFatura>
          <DataFatura >2026-08-13</DataFatura>
          <NifFatura >NNNNNNNNN</NifFatura>
          <ValorIVA >32.22</ValorIVA>
          <ValorTotalFatura >280.00</ValorTotalFatura>
          <Ficheiro  attachmentDate="2026-08-13T20:33:31Z" fileName="114" extension="pdf">BASE64</Ficheiro>
        </Quadro02T1-Linha>
      </Quadro02T1>
      <SomaIVA >32.22</SomaIVA>
      <SomaValorFatura >280.00</SomaValorFatura>
    </Quadro02>
  </Rosto>
</Simplex130>
```

## The four traps

1. **Element names are PascalCase in the file, camelCase in the XSD.** The
   schema says `numeroFatura`, `valorIVA`, `dataFatura`; the app writes
   `NumeroFatura`, `ValorIVA`, `DataFatura`. `DeclarationWriter.firstToUp()`
   does it on the way out. **A generator written from the XSD alone produces a
   file the app will not load**, which is the whole reason a saved sample was
   needed.
2. **`encoding='ISO-8859-1'`, single-quoted, and it is not decorative.**
   Supplier names carry Portuguese accents. Emit Latin-1 bytes, not UTF-8, or
   declare honestly and expect the app to disagree with the bytes.
3. **`SomaIVA` and `SomaValorFatura` are written into the file**, not computed on
   load. The XSD marks them `<formula>tquadro02T1.cvalorIVA</formula>`, but they
   are still persisted, so the generator must total the lines itself and must
   agree with them.
4. **Lines are numbered from 1** via `numero` on `Quadro02T1-Linha`, and the
   attribute is required.

5. **`TipoEntidade` is a single lowercase letter — `d`** for *Funcionários de
   Embaixadas e Organismos Internacionais*, which is the case this project needs.
   The string `Embaixada` appears **nowhere** in the jar or the app bundle, so the
   dropdown is filled from a catalog that cannot be read statically and the code
   can only be learned from a saved file. Do not try to derive it.

Also: the app writes a stray space before `>` on most tags (`<NumeroFatura >`,
`<Quadro02T1 >`, and two spaces on `<Ficheiro  `). That is the writer joining an
empty attribute map. Almost certainly cosmetic — but it is what the app itself
produces, so matching it costs nothing and rules out one class of surprise.

## The embedded document

`Ficheiro` holds the receipt **base64-encoded, inline, unwrapped**, with three
attributes:

| Attribute | Sample | Notes |
|---|---|---|
| `attachmentDate` | `2026-08-13T20:33:31Z` | `yyyy-MM-dd'T'HH:mm:ss'Z'` |
| `fileName` | `114` | **without** the extension |
| `extension` | `pdf` | separate |

**The app pads the blob with trailing NUL bytes.** In the sample, 172,776 bytes
decoded of which the PDF is 111,534 and the remaining 61,242 are zeros. A PDF is
self-terminating at `%%EOF`, so the padding looks like a buffer the handler never
truncates rather than anything meaningful. **Embed the exact file bytes and no
padding** — but that is the one assumption here that has not been proven by
loading a generated file back, so prove it before trusting a real submission.

## Mapping to the IVA section

| Simplex130 | HelpfulForms | Notes |
|---|---|---|
| `NumeroFatura` | `Número` | |
| `DataFatura` | `Date` | `yyyy-MM-dd` |
| `NifFatura` | `Emitente NIF` | |
| `ValorTotalFatura` | `Amount` | 2dp |
| `ValorIVA` | `IVA Amount` | 2dp |
| `Ficheiro` | the receipt in Drive | base64 |
| `CodigoBem` | `Tipo` | optional; absent in the sample |
| `Importado` | — | optional boolean; nothing maps to it yet |
| `NifEntidade` | **`REF_JALLC_NIF`** | the employing body, not you |
| `TipoEntidade` | — constant `d` | see trap 5 |
| `NifFuncionario` | **`REF_MY_NIF`** | you, the claimant |

**Quadro 01 is three fields, with an ordering dependency in the UI that does not
apply to a generated file.** On screen: `NifEntidade` first, then `TipoEntidade`
from the dropdown, and only once that is chosen does the **Funcionário** panel
appear to accept `NifFuncionario`. In the file they are three sibling elements
written at once, so the interaction is a UI affordance, not a format rule.

Both NIFs are already Script Properties — `REF_JALLC_NIF` and `REF_MY_NIF` — and
**neither belongs in this repo**, which is why every NIF above is `NNNNNNNNN`.

## `CodigoBem` — the per-line Tipo, and why its dropdown appears late

**The Quadro 02 `Tipo` dropdown is empty until `TipoEntidade` is chosen in Quadro
01**, which is the second interactive dependency in this form and the reason a
first save had no `CodigoBem`. It is not a quirk: `EntidadeEnum.getBensServicos()`
returns a *different allowed list per entity type*, so until the entity type is
known there is nothing to offer.

Read out of the jar rather than the UI, so it is complete rather than whatever
happened to be on screen. **For `d`, exactly eight codes are legal:**

| Code | Descrição |
|---|---|
| 101 | Vestuário e calçado |
| 102 | Electrodomésticos |
| 103 | Móveis |
| 104 | Jóias, bijutarias |
| 106 | Artigos de escritório |
| 107 | Outros bens não especificados |
| 156 | Reparação ou manutenção de veículos automóveis |
| 157 | Outros serviços não especificados |

**What `d` may NOT claim, though `c` — the embassy itself — may:** `105` bens
alimentares e bebidas, `151` trabalhos imobiliários, `152` água/gás/electricidade,
`153` serviços de alimentação e bebidas, `154` alojamento, `155` telefone. Food,
drink, utilities, accommodation and telephone are all closed to an employee. That
is a real eligibility rule, not a UI detail, and it is the kind of thing worth
knowing **before** filing a receipt rather than at submission.

The other entity types, for completeness: `a` Comunidades Religiosas → 301, 302;
`b` IPSS → 351, 352, 353, 358; `c` Embaixadas e Organismos Internacionais → all
of 101–107 and 151–157.

**This should drive the IVA category list in `Config.js`.** Those eight are the
only values the tax app will accept from this claimant, so any other category
guarantees a rejected line — an argument for the closed list being generated from
this table rather than typed.

## The cross-check worth keeping

AT's own types demand exactly what `VALIDATION-PLAN.md` already proposed, arrived
at independently from using the thing:

- `Decimal_15` is `fractionDigits="2" minInclusive="0"` — rules 1 and 2.
- `NifFatura` is `xs:long` — rule 3, nine digits and numeric only.

That is decent evidence those rules are right, and an argument for enforcing them
at intake rather than discovering them at submission.

## Before building the generator

- **Prove a generated file loads.** Round-trip the sample first: regenerate it
  from its own values and confirm DRORIVA opens it identically. Only then wire it
  to real rows.
- **Size.** Base64 inflates by a third and these are scans. One invoice here was
  111 KB of PDF for 231 KB of XML. Apps Script holds the whole thing in memory,
  so a large batch is the thing most likely to break first.
- **Which rows.** Presumably IVA entries that are complete and not yet submitted,
  which means a "submitted" marker the sheet does not have yet.
