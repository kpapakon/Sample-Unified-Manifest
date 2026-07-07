export class TaskpaneActions {
      static async enableButton(buttonId: string, enable: boolean) {
    await Office.ribbon.requestUpdate({
      tabs: [
        {
          id: 'TabHome',
          groups: [
            {
              id: 'CommandsGroup',
              controls: [
                {
                  id: buttonId,
                  enabled: enable
                }
              ]
            }
          ]
        }
      ]
    });
  }
}
