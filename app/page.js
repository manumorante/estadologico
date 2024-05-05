import styles from "@/css/home.module.scss"
import Link from "next/link"
export default function Home() {
  return (
    <>
      <div className={styles.info}>
        <div className="flexbox__item">
          <h2 className="home__title">
            Servicios de <strong>Diseño Web</strong> y{" "}
            <strong>Programación</strong>.
          </h2>

          <p className={styles.slogan}>Los tiempos cambian...</p>

          <Link
            href="/portafolio"
            title="Galería de trabajos realizados"
            className="button"
          >
            Ver portafolio
          </Link>
        </div>
      </div>

      <div className={styles.footer}>
        <p>
          {" "}
          &#8220;El diseño web no solo aporta a la comunicación textual
          (contenidos) existente en Internet una faceta visual, sino que obliga
          a pensar una mejor estructuración de los mismos en un nuevo soporte.
          La unión de un buen diseñ´ con una jerarquía bien elaborada de
          contenidos aumenta la eficiencia de la web como canal de comunicación
          e intercambio de datos, que brinda posibilidades como el contacto
          directo entre el productor y el consumidor de contenidos,
          característica destacable del medio Internet.&#8221;
        </p>
        &#8212; Wikipedia
      </div>
    </>
  )
}
