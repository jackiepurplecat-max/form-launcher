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

## The five traps

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
self-terminating at `%%EOF`, so the padding is a buffer the handler never
truncates rather than anything meaningful.

**Embed the exact file bytes and no padding. PROVEN 25 Aug**, and proven the only
way it could be: `v2/test/iva-roundtrip.js` wrote the sample back out with the
61,242 NUL bytes dropped — 149,747 bytes rather than 231,403 — and DRORIVA opened
it with every field intact and the receipt readable. So the padding is the app's
own and it does not expect it back. This was the last open assumption in the
format; nothing here is guesswork now.

## What may be attached — three rules the app does not enforce

From DRORIVA's own **Ajuda → Anexação de ficheiros**:

> - Não podem ser anexadas mais do que **50 faturas por pedido**;
> - Não são aceites ficheiros com mais do que **200 Kb**;
> - Apenas são aceites ficheiros anexos do formato **Portable Document Format**
>   (ficheiros com extensão pdf).
>
> Caso o pedido de restituição exceda as 50 faturas ou a dimensão do ficheiro
> exceda 200 Kb, admite-se a submissão de tantos pedidos quanto os necessários
> para completar o mesmo.

**None of the three is in the XSD and none is in the jar** — grep it, the help
text is not there either. So the app does not stop you: a 3 MB JPEG attaches,
saves, and reloads fine, and the refusal arrives from AT at submission. That
makes these intake rules rather than format rules, and it is why they are
enforced in this project rather than trusted to the app.

Where they live now:

| Rule | Config | Enforced |
|---|---|---|
| PDF only | `iva.documentRules.extensions` | `documentRuleProblem()` — upload, pick, page, and again at export |
| 200 Kb per file | `iva.documentRules.maxBytes` | same |
| 50 faturas per pedido | `iva.declaration.maxLinesPerDeclaration` | `ivaAttachDocuments()` — deferred to the next batch |

**A broken file and a full batch are different things**, and the export keeps
them apart. Over 50 lines is AT's own "submit as many pedidos as needed", so the
51st is *deferred*: still `To Do`, reported, exported next run. A non-PDF or an
oversized scan is a fault in that row — it will be just as broken next time — so
it is *rejected*, named, and **does not stop the rows behind it**.

**200 Kb is read as per attached file, not per declaration.** The bullet sits in
a list about `ficheiros anexos`, alongside the PDF rule, which is the strongest
reading — but the closing sentence says "a dimensão do ficheiro" in the singular
and could just be loose wording for the whole submission. If a real submission
is ever refused on size while every attachment is under 200 Kb, that is the
answer, and the fix is a total-bytes cap beside `maxBase64Bytes`.

## Mapping to the IVA section

| Simplex130 | HelpfulForms | Notes |
|---|---|---|
| `NumeroFatura` | `Número` | |
| `DataFatura` | `Date` | `yyyy-MM-dd` |
| `NifFatura` | `Emitente NIF` | |
| `ValorTotalFatura` | `Amount` | 2dp |
| `ValorIVA` | `IVA Amount` | 2dp |
| `Ficheiro` | the receipt in Drive | base64 |
| `CodigoBem` | `Tipo` | optional in the schema; **present in the sample** as `101` |
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

## The Submit button — what it actually does

Read out of the jar on 25 Aug, because "can the export press Submit too?" is the
obvious next question. **Nothing is built and nothing should be until the
caveats at the end are answered.** This is what the button is, not a plan.

`Simplex130SubmitAction` → `PortalTaxClientRevampedSubmitAction` →
`SubmitTaxHTTPConnectionRevamped`. It is **one form POST**, not SOAP:

```
POST https://www.portaldasfinancas.gov.pt/pt/externalws/drorivaws/entregaDeclORIVAOffline.action
```

| Parameter | Value |
|---|---|
| `systemName` | `taxUser` — constant |
| `systemPass` | `Password.encryptWithPaddingLength("y5jeC$maQe")` — a shared secret compiled into the app |
| `decl` | `URLEncoder.encode(Base64(Deflater(xml)), "UTF-8")` — **the same XML this repo already generates** |
| `SPA`, `nifA`, `passwordA` | your Portal das Finanças NIF and password |
| `SPB`/`nifB`/`passwordB`, `TOC`/`toc`/`passwordToc` | second taxpayer and accountant — not used here |
| `app_version` | `Simplex130Parameters.getAppVersion()` |
| `checkwarnings` | boolean |
| `ws_version` | `1.0` |

The reply is XML: `<erros>` with `<error>` children carrying `code`, `field`,
`severity` and a **base64** `msg`, or a `receipt` — the comprovativo, also
base64. So a real submission returns a receipt worth filing beside the rows.

**The declaration is the only hard part already solved.** `decl` is
`ivaDeclarationXml()` output, deflated and base64'd, so the format work carries
straight over.

**Two things stand between here and a working `uiIvaSubmit()`:**

1. **`ZipUtil.deflate` is `java.util.zip.DeflaterOutputStream`** — a raw zlib
   stream, not a ZIP archive. Apps Script has `Utilities.zip`, which produces an
   archive, and `Utilities.gzip`, which adds a gzip header — **neither is what
   the endpoint expects**. It needs a deflate implementation in the script, or a
   stored-block encoder, which is legal deflate and about thirty lines.
2. **`Password.encryptWithPaddingLength`** pads the secret to 16 characters with
   a random prefix, then encrypts it with itself as the key via `pt.opensoft.crypt.Crypt`.
   Reimplementable, but it has to be exact and it is the one part with no
   round-trip test available offline.

**And three reasons to think twice before doing it:**

- **There is no test endpoint.** `url.submit.test` points at
  `portalfin.stinternetdev.ritta.local`, an internal host. Every attempt is a
  live filing against a tax authority — no dry run exists.
- **It puts a Portal das Finanças password in Script Properties.** The NIFs are
  already there; a password that files tax declarations is a different class of
  secret, on an Apps Script project reached through a web app.
- **The app is doing something for you that is not in this list** — retries,
  session handling, the warnings round trip (`checkwarnings` implies a
  submit-twice flow: once to collect warnings, once to confirm). Matching the
  parameters is not the same as matching the behaviour.

**The honest middle option** is what already exists plus one step: generate the
declaration, open DRORIVA, `File → Abrir`, press Submit. The retyping — which is
what the export was for — is gone either way. Automating the last click buys
one click and takes on all of the above.

## Before building the generator

- **Prove a generated file loads.** Round-trip the sample first: regenerate it
  from its own values and confirm DRORIVA opens it identically. Only then wire it
  to real rows.
- **Size.** Base64 inflates by a third and these are scans. One invoice here was
  111 KB of PDF for 231 KB of XML. Apps Script holds the whole thing in memory,
  so a large batch is the thing most likely to break first.
- **Which rows.** Presumably IVA entries that are complete and not yet submitted,
  which means a "submitted" marker the sheet does not have yet.
