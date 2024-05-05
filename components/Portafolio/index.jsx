import styles from "./styles.module.scss"
import Link from "next/link"
import Image from "next/image"

export default function PortafolioItem({
  image,
  title,
  url = "#",
  description = "",
}) {
  return (
    <div class={styles.item}>
      <Link href={url} class={styles.image}>
        <Image src={`/img/webs/${image}`} width={115} height={85} alt={title} />
      </Link>

      <Link href={url} className={styles.title}>
        {title}
      </Link>

      <p class={styles.description}>{description}</p>
    </div>
  )
}
