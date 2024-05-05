import styles from "./styles.module.scss"
import Item from "@/components/Portafolio"

export default function Home() {
  return (
    <>
      <div className={styles.portfolio}>
        <div className={styles.items}>
          <Item
            title="Fernsehschoner Project"
            url="https://www.fernsehschoner.de"
            image="fersehschoner.jpg"
            description="Sitio Flash para Gregor Kuschmir, media-artist de la empresa de creativos La Fábrica."
          />
          <Item
            title="Trade Punk"
            url="https://www.tradepunk.com"
            image="tradepunk.jpg"
            description="Web oficial del grupo Trade Punk. Granada."
          />
          <Item
            title="La Red Social"
            url="https://www.laredsocial.org"
            image="laredsocial.jpg"
            description="Ayuda y tutoriales en video sobre redes sociales, Facebook, Tuenti, MySpace, Windows Live, ..."
          />
          <Item
            title="Formas Formación"
            image="formasformacion.jpg"
            description="Formación presencial y distancia. Prácticas en Centros Concertados y Bolsa de Trabajo Activa."
          />
          <Item
            title="Una Vez Fuí"
            url="#"
            image="unavezfui.jpg"
            description="Ayuda, videos y noticias sobre cosicas modernas de internet."
          />
        </div>
      </div>
    </>
  )
}
