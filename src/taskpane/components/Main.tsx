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

        setState({ ...state, cells: range.values, error: "" });
      });
    } catch (error) {
      console.error(error);
    }
  };

  const handlePaste = async () => {
    try {
      if (!state.cells.length) {
        setState({ ...state, error: "You should copy data first" });
      } else {
        await Excel.run(async context => {
          const range = context.workbook.getSelectedRange();
          range.load("address");
          range.load("values");
          await context.sync();

          let rows = state.cells.length;
          let cols = state.cells[0].length;
          let range2 = range.getAbsoluteResizedRange(rows, cols);

          range2.values = state.cells;
          await context.sync();
        });
      }
    } catch (error) {
      setState({ ...state, error: JSON.stringify(error) });
      console.error(error);
    }
  };

  return (
    <div className="container">
      {state.error.length ? <p className="error">{state.error} </p> : null}
      <Button className="button" buttonType={ButtonType.primary} onClick={handleCopy}>
        COPY
      </Button>

      <Button className="button" buttonType={ButtonType.primary} onClick={handlePaste}>
        PASTE
      </Button>
    </div>
  );
};

export default Component;
