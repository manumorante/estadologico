import styles from "./styles.module.scss"
import Link from "next/link"
import Image from "next/image"

export default function Home() {
  return (
    <>
      <div className={styles.info}>
        <div className={styles.text}>
          <h2>
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

        <Image
          className={styles.cover}
          src="/img/tranvia.jpg"
          width={1800}
          height={836}
          alt="Tranvía foto por Manu Morante"
        />
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
