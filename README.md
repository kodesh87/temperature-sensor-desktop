# Temperature monitoring — serial acquisition client (Visual Basic 6)

Reads a temperature sensor over a serial port and forwards each reading to a
remote web endpoint over raw HTTP. This is the acquisition half of a two-part
system; the receiving half is
[kodesh87/temperature-sensor-web](https://github.com/kodesh87/temperature-sensor-web).

> **University coursework, 2010.** Published as a record of early work. Visual
> Basic 6 reached end of support long ago and nothing here reflects how I build
> software now — see [Known limits](#known-limits).

## How it works

<img width="873" alt="System architecture: serial sensor to VB6 client to PHP endpoints to MySQL, with a polling browser display" src="https://user-images.githubusercontent.com/54540612/172774474-c22cd356-a533-49ab-8dea-ccd021300b10.png">

```
sensor ──serial──> MSComm ──> this app ──raw HTTP GET via Winsock──> update.php ──> MySQL
```

**Serial acquisition.** `MSComm1` is configured with `RThreshold = 1`, so the
`OnComm` event fires on every received character rather than on a polling timer.
Readings arrive as they are produced instead of being sampled.

**Transport.** There is no HTTP client — VB6 has none built in — so the request
line and headers are assembled by hand as a string and written to a `Winsock`
socket. `url.bas` is a small parser that splits a URL into scheme, host, port,
URI, and query so the socket can be opened against the right endpoint.

**Two alternating sockets.** `Winsock1` and `Winsock2` are used in turn, tracked
by `winsockUsed`. A Winsock control does not become available again the instant
`Close` is called, so at a one-second send interval a single control can still be
closing when the next reading is ready. Alternating between two avoids dropping
that reading. It is a workaround for a platform constraint, not a design
preference.

## Layout

```
Form1.frm       UI, MSComm handling, Winsock send logic
url.bas         URL parser — scheme / host / port / URI / query
Project1.vbp    VB6 project file
Package/        Packaging & Deployment Wizard output — CAB, setup, VB6 runtime
```

`Package/Source Code WEB/` contains a copy of the web half as it was bundled for
submission. The maintained version is in the web repository linked above.

## Running it

Requires a 32-bit Windows environment with the Visual Basic 6 runtime, plus
`MSCOMM32.OCX` and `MSWINSCK.OCX` registered. Both are included under
`Package/Support/`. The target endpoint is set in `Form1.frm`; it currently points
at hosts that no longer exist.

## Known limits

- **Endpoint URLs are hardcoded** in the form rather than read from configuration,
  and the hosts they name are long gone.
- **HTTP is hand-rolled** over a raw socket — no status handling, no timeout, no
  retry. A failed send is a lost reading.
- **No authentication** on the receiving endpoint, so any client could post
  readings. See the web repository for the matching note.
- **Alternating sockets** is a workaround, not a solution. A send queue would be
  the correct fix.
- **Compiled executables are committed** alongside the source.

## Context

Built in 2010 as coursework, paired with a PHP web application that stored and
displayed the readings. Keeping both halves published makes the system legible as
a whole rather than as two orphaned fragments.

Current work — enterprise integration, identity and single sign-on, and backend
systems — is in private client and employer repositories. Professional history:
[linkedin.com/in/kodesh87](https://linkedin.com/in/kodesh87).
