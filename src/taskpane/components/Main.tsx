import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";

const Component = () => {
  const [state, setState] = React.useState({ value: "" });

  return (
    <div className="container">
      <p>{state.value}</p>
      {state.value.length ? <p>{state.value} </p> : null}
      <Button
        className="button"
        buttonType={ButtonType.primary}
        onClick={() => setState({ ...state, value: "monther" })}
      >
        COPY
      </Button>

      <Button
        className="button"
        buttonType={ButtonType.primary}
        onClick={() => setState({ ...state, value: "monther" })}
      >
        PASTE
      </Button>
    </div>
  );
};

export default Component;
