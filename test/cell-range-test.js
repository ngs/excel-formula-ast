const { buildTree, stringify } = require('../');
const { tokenize } = require('@modelmap/excel-formula-tokenizer');
const { deepStrictEqual } = require('assert');
const builder = require('../lib/node-builder');

describe('cell ranges', function() {
  it('A1', function() {
    const tree = buildTree(tokenize('A1'));

    deepStrictEqual(tree, builder.cell('A1', 'relative'));
    deepStrictEqual(stringify(tree), 'A1');
  });

  it('A$1', function() {
    const tree = buildTree(tokenize('A$1'));

    deepStrictEqual(tree, builder.cell('A$1', 'mixed'));
    deepStrictEqual(stringify(tree), 'A$1');
  });

  it('$A1', function() {
    const tree = buildTree(tokenize('$A1'));

    deepStrictEqual(tree, builder.cell('$A1', 'mixed'));
    deepStrictEqual(stringify(tree), '$A1');
  });

  it('$A$1', function() {
    const tree = buildTree(tokenize('$A$1'));

    deepStrictEqual(tree, builder.cell('$A$1', 'absolute'));
    deepStrictEqual(stringify(tree), '$A$1');
  });

  it('A1:A4', function() {
    const tree = buildTree(tokenize('A1:A4'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('A1', 'relative'), builder.cell('A4', 'relative')));
    deepStrictEqual(stringify(tree), 'A1:A4');
  });

  it("'Foo ! Bar'!A1:A4", function() {
    const tree = buildTree(tokenize("'Foo ! Bar'!A1:A4"));

    deepStrictEqual(tree, builder.cellRange(builder.cell('Foo ! Bar!A1'), builder.cell('Foo ! Bar!A4')));
    deepStrictEqual(stringify(tree), "'Foo ! Bar'!A1:A4");
  });

  it("'Foo ! Bar'!A1:'Foo ! Bar'!A4", function() {
    const tree = buildTree(tokenize("'Foo ! Bar'!A1:'Foo ! Bar'!A4"));

    deepStrictEqual(tree, builder.cellRange(builder.cell('Foo ! Bar!A1'), builder.cell('Foo ! Bar!A4')));
    deepStrictEqual(stringify(tree), "'Foo ! Bar'!A1:'Foo ! Bar'!A4");
  });

  it('$A1:A$4', function() {
    const tree = buildTree(tokenize('$A1:A$4'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('$A1', 'mixed'), builder.cell('A$4', 'mixed')));
    deepStrictEqual(stringify(tree), '$A1:A$4');
  });

  it('$A$1:$A$4', function() {
    const tree = buildTree(tokenize('$A$1:$A$4'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('$A$1', 'absolute'), builder.cell('$A$4', 'absolute')));
    deepStrictEqual(stringify(tree), '$A$1:$A$4');
  });

  it('1:4', function() {
    const tree = buildTree(tokenize('1:4'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('1', 'relative'), builder.cell('4', 'relative')));
    deepStrictEqual(stringify(tree), '1:4');
  });

  it('$1:4', function() {
    const tree = buildTree(tokenize('$1:4'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('$1', 'absolute'), builder.cell('4', 'relative')));
    deepStrictEqual(stringify(tree), '$1:4');
  });

  it('C:G', function() {
    const tree = buildTree(tokenize('C:G'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('C', 'relative'), builder.cell('G', 'relative')));
    deepStrictEqual(stringify(tree), 'C:G');
  });

  it('C:$G', function() {
    const tree = buildTree(tokenize('C:$G'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('C', 'relative'), builder.cell('$G', 'absolute')));
    deepStrictEqual(stringify(tree), 'C:$G');
  });

  it('C:G5', function() {
    const tree = buildTree(tokenize('C:G5'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('C', 'relative'), builder.cell('G5', 'relative')));
    deepStrictEqual(stringify(tree), 'C:G5');
  });

  it('5:D5', function() {
    const tree = buildTree(tokenize('5:D5'));

    deepStrictEqual(tree, builder.cellRange(builder.cell('5', 'relative'), builder.cell('D5', 'relative')));
    deepStrictEqual(stringify(tree), '5:D5');
  });

  it('A1:B3,C1:D3', function() {
    const tree = buildTree(tokenize('A1:B3,C1:D3'));

    deepStrictEqual(
      tree,
      builder.binaryExpression(
        ',',
        builder.cellRange(builder.cell('A1', 'relative'), builder.cell('B3', 'relative')),
        builder.cellRange(builder.cell('C1', 'relative'), builder.cell('D3', 'relative')),
      ),
    );
    deepStrictEqual(stringify(tree), 'A1:B3,C1:D3');
  });

  it('A1:B3 B1:D3', function() {
    const tree = buildTree(tokenize('A1:B3 B1:D3'));

    deepStrictEqual(
      tree,
      builder.binaryExpression(
        ' ',
        builder.cellRange(builder.cell('A1', 'relative'), builder.cell('B3', 'relative')),
        builder.cellRange(builder.cell('B1', 'relative'), builder.cell('D3', 'relative')),
      ),
    );
    deepStrictEqual(stringify(tree), 'A1:B3 B1:D3');
  });
});
