import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { NavCustomLinkExample } from "./Navigation";
import { Nav, INavStyles, INavLinkGroup, INavLink } from "@fluentui/react/lib/Nav";
import NavComponent from "./Navigation";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

/* nav   */

const navStyles: Partial<INavStyles> = {
  root: {
    width: 208,
    height: 350,
    boxSizing: "border-box",
    border: "1px solid #eee",
    overflowY: "auto",
  },
};

const navLinkGroups: INavLinkGroup[] = [
  {
    links: [
      {
        name: "Start",
        url: "taskpane.html",
        icon: "News",
        key: "key1",
        target: "_blank",
      },
      {
        name: "Aufträge",
        url: "http://example.com",
        expandAriaLabel: "Expand Auftrag section",
        links: [
          {
            name: "Erstellen",
            url: "http://msn.com",
            disabled: true,
            key: "key2",
            target: "_blank",
          },
          {
            name: "Verwalten",
            url: "http://msn.com",
            disabled: true,
            key: "key3",
            target: "_blank",
          },
        ],
        isExpanded: false,
      },
      {
        name: "Mitarbeiter",
        url: "http://example.com",
        key: "key4",
        isExpanded: true,
        target: "_blank",
      },
      {
        name: "Arbeitsplanung",
        url: "http://example.com",
        key: "key5",
        target: "_blank",
      },
    ],
  },
];

export const NavBasicExample: React.FunctionComponent = () => {
  return (
    <Nav
      onLinkClick={_onLinkClick}
      selectedKey="key1"
      ariaLabel="Nav basic example"
      styles={navStyles}
      groups={navLinkGroups}
      isOnTop

    />
  );
};

function _onLinkClick(_ev?: React.MouseEvent<HTMLElement>, item?: INavLink) {
  if (item && item.name === "News") {
    alert("News link clicked");
  }
}

/* nav   */

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
        range.format.fill.color = "#66ff00";
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
        range.format.fill.color = "#66ff00";
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
        range.format.fill.color = "#66ff00";
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
        <NavBasicExample></NavBasicExample>
    return (
      <div className="ms-welcome">
        <Header
          logo={require("./../../../assets/logo-filled.png")}
          title={this.props.title}
          message="Herzlich Willkommen"
        />        <NavBasicExample></NavBasicExample>

        <HeroList
          message="One-Klick Tabellen-, Spalten-, Zeilen- und Zellenformatierung zur vereinfachung und standardisierung der Arbeitplannung"
          items={this.state.listItems}
        >        <NavBasicExample></NavBasicExample>

          <p className="ms-font-xs">
            Markiere den gewünschten <b>Bereich</b> oder eine einzelne Zelle und klicke auf die gewünschte
            <b>Formatierungsvorlage</b>.
          </p>

          <h3>Status</h3>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm4 ms-smPush8">First in code</div>
            <div className="ms-Grid-col ms-sm8 ms-smPull4">Second in code</div>
          </div>
          <DefaultButton
            color="#66ff00"
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickfertig}
          >
            Fertig
          </DefaultButton>
          <DefaultButton
            color="#66ff00"
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickläuft}
          >
            Läuft
          </DefaultButton>
          <DefaultButton
            color="#66ff00"
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickläuftbald}
          >
            Läuft bald
          </DefaultButton>
          <DefaultButton
            color="yellow"
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickpausiert}
          >
            Pausiert
          </DefaultButton>
          <DefaultButton
            color="yellow"
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.clickfreigabewartend}
          >
            warten auf Freigabe
          </DefaultButton>
          {/* 
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm4 ms-smPush8">
                <div className="LayoutPage-demoBlock">First in code</div>
              </div>
              <div className="ms-Grid-col ms-sm8 ms-smPull4">
                <div className="LayoutPage-demoBlock">Second in code</div>
              </div>
            </div>
          </div>
          <p className="ms-font-l">
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
