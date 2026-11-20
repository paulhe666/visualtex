# Equations written by MathType itself

Each `.native` file is the `Equation Native` stream (header + MTEF 5) of a MathType OLE object
saved by MathType, not by VisualTeX. `MathTypeRealCorpusTests` reads every one of them.

Source: https://github.com/jure/mathtype, `spec/fixtures/input/mathtype5/*.bin`
(commit 80413caa2b57d957a3c4ec20c2a681606a937c92), with the `Equation Native` stream taken
out of each compound file. The same repository's `spec/fixtures/expected/mathtype5/*.xml`
is an independent parse of the same files.

## License of the fixtures

Copyright (c) 2015 Jure Triglav

MIT License

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
