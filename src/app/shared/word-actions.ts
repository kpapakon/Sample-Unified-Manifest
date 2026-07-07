export class WordActions {
  static async sayHello(color: string, action?: boolean) {
    await Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph(
        action ? 'This auto-open taskpane is ON' : 'This auto-open taskpane is OFF',
        Word.InsertLocation.end,
      );
      paragraph.font.color = color;
      await context.sync();
    });
  }

  static async underlineSelection() {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (selection === null) return;

      selection.font.underline = Word.UnderlineType.single;
      await context.sync();
    });
  }

  static async getTextSelection(document: Office.Document): Promise<string> {
    return new Promise((resolve, reject) => {
      document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const value = result.value;
          if (typeof value === 'string') {
            resolve(value);
          } else {
            reject(new Error('Selected data is not a string.'));
          }
        } else {
          reject(
            new Error(`Error getting selected data: ${result.error.message}`),
          );
        }
      });
    });
  }
}
