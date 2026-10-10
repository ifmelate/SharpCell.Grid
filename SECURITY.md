# Security

SharpCell.Grid shows workbooks that often come from untrusted sources. A file that makes it crash
the process, hang or use unbounded memory is a security issue; one inside the formula engine or
the .xlsx reader belongs to [SharpCell](https://github.com/ifmelate/SharpCell/security).

## Reporting

Report privately through GitHub: **Security → Report a vulnerability** on
https://github.com/ifmelate/SharpCell.Grid. Do not open a public issue. Attach the smallest file
that shows the problem, and do not send confidential workbooks.

The project has one maintainer working in spare time. Expect an answer within two weeks; a fix
ships in a new release of the latest minor version.

## Supported versions

Only the latest released version gets fixes.
