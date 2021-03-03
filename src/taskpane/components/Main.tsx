import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";

const Component = () => {
  const [state, setState] = React.useState({ cells: [], error: "" });

  const handleCopy = async () => {
    try {
      await Excel.run(async context => {
        const range = context.workbook.getSelectedRange();

        range.load("address");
        range.load("values");

        await context.sync();

        setState({ ...state, cells: range.values });
      });
    } catch (error) {
      console.error(error);
    }
  };

  return (
    <div className="container">
      <p>{state.cells}</p>
      <Button className="button" buttonType={ButtonType.primary} onClick={handleCopy}>
        COPY
      </Button>

      <Button className="button" buttonType={ButtonType.primary} onClick={() => setState({ ...state })}>
        PASTE
      </Button>
    </div>
  );
};

export default Component;
