import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Steigerung der Betriebsqualität",
        },
        {
          icon: "Unlock",
          primaryText: "Freigeben von Unternehmensressourcen",
        },
        {
          icon: "Design",
          primaryText: "Förderung einheitlicher Firmenkultur und -identität",
        },
      ],
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
        // Update the fill color
        range.format.fill.color = "yellow";
        range.format.font.color = "red";
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  clickfertig = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const data = [["fertig"]];
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
        range.values = data;
        // Update the fill color
        range.format.fill.color = "green";
        // Update the Text color
        range.format.font.color = "black";
        // Update the Schrifttyp zu "Fett"
        range.format.font.bold = true;
        // Spaltenbreite automatisch anpassen
        range.format.autofitColumns();
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  clickläuft = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const data = [["läuft"]];
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
        range.values = data;
        // Update the fill color
        range.format.fill.color = "green";
        // Update the Text color
        range.format.font.color = "black";
        // Update the Schrifttyp zu "Fett"
        range.format.font.bold = true;
        // Spaltenbreite automatisch anpassen
        range.format.autofitColumns();
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  clickläuftbald = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const data = [["Läuft bald"]];
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
        range.values = data;
        // Update the fill color
        range.format.fill.color = "green";
        // Update the Text color
        range.format.font.color = "black";
        // Update the Schrifttyp zu "Fett"
        range.format.font.bold = true;
        // Spaltenbreite automatisch anpassen
        range.format.autofitColumns();
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  clickpausiert = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const data = [["Pausiert"]];
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
        range.values = data;
        // Update the fill color
        range.format.fill.color = "yellow";
        // Update the Text color
        range.format.font.color = "black";
        // Update the Schrifttyp zu "Fett"
        range.format.font.bold = true;
        // Spaltenbreite automatisch anpassen
        range.format.autofitColumns();
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  clickfreigabewartend = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const data = [["warten auf Freigabe"]];
        const range = context.workbook.getSelectedRange();
        // Read the range address
        range.load("address");
        range.values = data;
        // Update the fill color
        range.format.fill.color = "yellow";
        // Update the Text color
        range.format.font.color = "black";
        // Update the Schrifttyp zu "Fett"
        range.format.font.bold = true;
        // Spaltenbreite automatisch anpassen
        range.format.autofitColumns();
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Initialisierung Fehlgeschlagen - Body leer - sideload?"
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          logo={require("./../../../assets/logo-filled.png")}
          title={this.props.title}
          message="Herzlich Willkommen"
        />
        <HeroList
          message="One-Klick Tabellen-, Spalten-, Zeilen- und Zellenformatierung zur vereinfachung und standardisierung der Arbeitplannung"
          items={this.state.listItems}
        >
          <p className="ms-font-xs">
            Markiere den gewünschten <b>Bereich</b> oder eine einzelne Zelle und klicke auf die gewünschte
            <b>Formatierungsvorlage</b>.
          </p>
          <h2>Offene Kapselaufträge</h2>
          <h3>Status</h3>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickfertig}
          >
            Fertig
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickläuft}
          >
            Läuft
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickläuftbald}
          >
            Läuft bald
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickpausiert}
          >
            Pausiert
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickfreigabewartend}
          >
            warten auf Freigabe
          </DefaultButton>
          {/* <p className="ms-font-l">
            Klicke auf eine <b>Aufgabe</b> um zu starten.
          </p>
          <h2>Aufträge</h2>
          <h3>Dosenaufträge</h3>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Dosenauftrag erfassen (AFK-_)
          </DefaultButton>
          <h3>Kapselaufträge</h3>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Kapselauftrag erfassen (_-A)
          </DefaultButton>
          <h3>Liquidaufträge</h3>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Liquidauftrag erfassen (ALF-_)
          </DefaultButton>
          <h3>Bulk- Doypackaufträge</h3>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Bulk- Doypackauftrag erfassen (AF-_-_)
          </DefaultButton> */}
        </HeroList>
      </div>
    );
  }
}
/**
 * Vorlage
 * https://forms.office.com/Pages/ShareFormPage.aspx?id=xha_d6UfpkaQSaRBx64oPOHJw-VAVsJCnlaVMUqVOSFUOE1IVzVEMFE1Nk9UT0dPOUY5S0RPQjhEUiQlQCN0PWcu&sharetoken=o57VGgQtOi3HP6XdDLYH
 * Für Zusammenarbeit
 * https://forms.office.com/Pages/DesignPage.aspx?fragment=FormId%3Dxha_d6UfpkaQSaRBx64oPOHJw-VAVsJCnlaVMUqVOSFUOE1IVzVEMFE1Nk9UT0dPOUY5S0RPQjhEUiQlQCN0PWcu%26Token%3Da7b7263b02cb4322bca2e2903b516e05
 * Antworten senden und sammeln
 * https://forms.office.com/Pages/ResponsePage.aspx?id=xha_d6UfpkaQSaRBx64oPOHJw-VAVsJCnlaVMUqVOSFUOE1IVzVEMFE1Nk9UT0dPOUY5S0RPQjhEUiQlQCN0PWcu
 * https://forms.office.com/e/VqLtWXKmY7
 * <iframe width="640px" height= "480px" src= "https://forms.office.com/Pages/ResponsePage.aspx?id=xha_d6UfpkaQSaRBx64oPOHJw-VAVsJCnlaVMUqVOSFUOE1IVzVEMFE1Nk9UT0dPOUY5S0RPQjhEUiQlQCN0PWcu&embed=true" frameborder= "0" marginwidth= "0" marginheight= "0" style= "border: none; max-width:100%; max-height:100vh" allowfullscreen webkitallowfullscreen mozallowfullscreen msallowfullscreen> </iframe>
 */
