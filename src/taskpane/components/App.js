import * as React from "react";
import PropTypes from "prop-types";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import getDocumentAsPdf from "../../commands/doc-to-pdf";
/* global Word */

const listItems = [
  {
    icon: "Ribbon",
    primaryText: "Achieve more with Office integration",
  },
  {
    icon: "Unlock",
    primaryText: "Unlock features and functionality",
  },
  {
    icon: "Design",
    primaryText: "Create and visualize like a pro",
  },
];

function App(props) {
  const { title, isOfficeInitialized } = props;

  const populateBasicText = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Test paragraph", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo="assets/logo-filled.png" title={title} message="Test Add-in" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={populateBasicText}
        >
          Populate doc with basic text
        </Button>
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={getDocumentAsPdf}
        >
          Turn doc into PDF
        </Button>
      </HeroList>
    </div>
  );
}

export default App;

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
