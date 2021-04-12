const alpha = [ 'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O',
  'P','Q','R','S','T','U','V','W','X','Y','Z' ];

interface rc {
  r: number;
  c: number;
}

export class Address {
  public static calcRelativePos (origPos: rc, targPos: rc) {
    let relativePos: rc;
    relativePos.r = targPos.r - origPos.r;
    relativePos.c = targPos.c - origPos.c;
    return relativePos;
  }

  public static calcTargetPos (origPos: rc, relativePos: rc) {
    let targPos: rc;
    targPos.r = origPos.r + relativePos.r;
    targPos.c = origPos.c + relativePos.c;
    return targPos;
  }
  public static rc2address (pos: rc): string {
    let col = "";
    let colNumber = pos.c;
    if (colNumber < 1 || colNumber > 16384) {
      throw new Error(`${colNumber} is out of bounds.`);
    }
    let x = colNumber;
    let y = 0;
    while (x > 26) {
      y = x % 26;
      x = Math.floor(x / 26);
      col = alpha[y - 1] + col;
    }
    col = alpha[x - 1] + col;
    return col + pos.r;
  }

  public static address2rc (address: string): rc {
    let hasCol = false;
    let colNumber = 0;
    let hasRow = false;
    let rowNumber = 0;
    for (let i = 0, char; i < address.length; i++) {
      char = address.charCodeAt(i);
      if (!hasRow && char >= 65 && char <= 90) {
        // 65 = 'A'.charCodeAt(0)
        // 90 = 'Z'.charCodeAt(0)
        hasCol = true;
        // colNumber starts from 1
        colNumber = (colNumber * 26) + char - 64;
      } else if (char >= 48 && char <= 57) {
        // 48 = '0'.charCodeAt(0)
        // 57 = '9'.charCodeAt(0)
        hasRow = true;
        // rowNumber starts from 0
        rowNumber = (rowNumber * 10) + char - 48;
      } else if (hasRow && hasCol && char !== 36) {
        // 36 = '$'.charCodeAt(0)
        break;
      }
    }
    if (!hasCol) {
      colNumber = undefined;
    } else if (colNumber > 16384) {
      throw new Error(`Out of bounds. Invalid column letter: ${colNumber}`);
    }
    if (!hasRow) {
      rowNumber = undefined;
    }
    return {r: rowNumber, c: colNumber};
  }
}
