import * as React from "react";
import Header from "./Header";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}


export interface AppState {
  listItems: any[];
  loading: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      loading: true
    };
  }

  componentDidMount() {
    setTimeout(() => {

      this.setState({
        loading: false,
        listItems: [
          {
            title: "Acerca de Plain Concepts",
            text: `Plain Concepts es una compañía especializada en tecnologías Microsoft, metodologías ágiles, gestión del ciclo de vida de aplicaciones, tuning de rendimiento, depuración avanzada en entornos, arquitectura de software, UX.
            Plain Concepts se centra en la consultoría de alto nivel y especialización, mentoring y Formación.
            Nuestro espectacular equipo de expertos, compuesto por varios Microsoft Valuable Professionals y Microsoft Certified Trainers, es crucial para conseguir nuestros resultados de satisfacción con los clientes. Todos nuestros profesionales son miembros de distintas comunidades técnicas y suelen participar como speakers en multitud de eventos a nivel mundial y son autores de libros, cursos y artículos públicos.
            `
          },
          {
            title: "Validez de la propuesta",
            text: `El período de validez de esta propuesta es de 5 días naturales a partir de la fecha de entrega de la misma. Excepcionalmente, en caso de darse aceptación de la propuesta por ambas partes tras dicho plazo, las fechas de calendario podrían retrasarse.`
          },
          {
            title: "Limitación de responsabilidad",
            text: `Queda excluida cualquier responsabilidad de Plain Concepts por las pérdidas indirectas o consecuenciales, entendiéndose por tales, otras perdidas distintas de los daños directos a los que se hace referencia en el apartado anterior, incluidos el lucro cesante y los costos en los que se incurra para prevenir o determinar las pérdidas consecuenciales.`
          },
          {
            title: "Indemnización máxima",
            text: `La responsabilidad de Plain Concepts respecto del cliente que surja de este contrato está limitada y no excederá de la cantidad pactada como remuneración total para este contrato, excepto en el caso que medie dolo, negligencia o mala fe por parte de Plain Concepts.
            La responsabilidad de Plain Concepts no superará en ningún caso el 100 % de los importes totales facturados en virtud del contrato de proyecto afectado y de los importes aún no facturados por los servicios ya proporcionados`
          },
          {
            title: "Premios y certificaciones",
            text: `• Gold Application Development
            • Gold Application Lifecycle Management
            • Gold Cloud Platform
            • Silver Application Integration
            • Silver Collaboration `
          },
          {
            title: "Premios internacionales",
            text: `• ALM Partner of the Year durante 6 años consecutivos
            • Windows 8 Applications Partner of the Year
            • FWA: Project Prometheus Training Center`
          }
        ]
      });
    }, 2000);
  }

  click = async (item) => {
    return Word.run(async context => {
      const heading = context.document.body.insertParagraph(item.title, Word.InsertLocation.end);
      heading.styleBuiltIn = "Heading1";

      const parts: string[] = item.text.split(`\n           `);
      parts.forEach(text => {
        
      const p = context.document.body.insertParagraph(text, Word.InsertLocation.end);
      p.styleBuiltIn = "Normal";
      });
      await context.sync();
    });
  };

  render() {
    // const { title, isOfficeInitialized } = this.props;

    // if (!isOfficeInitialized) {
    //   return (
    //     <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    //   );
    // }


    return (
      <div className="ms-welcome">
        <Header logo="assets/logo.png" title={this.props.title} message="Apartados para el documento" />
        {this.state.loading && <Spinner className="pc-spinner" size={SpinnerSize.large} label="Cargando plantillas corporativas..."/>}
        {!this.state.loading && <div className="pc-item-list ms-font-m ms-fontColor-neutralPrimary">
          {this.state.listItems.map(item => <div onClick={() => this.click(item)} className="pc-item">
              <div className="pc-item-title">{item.title}</div>
              <div className="pc-item-text">{item.text.substr(0, 100)}...</div>
            </div>)}
        </div>}
      </div>
    );
  }
}
