import styles from "./styles.module.scss"
export default function Footer() {
  const currentYear = new Date().getFullYear()
  return (
    <footer class={styles.footer}>
      <small>&copy; 2001-{currentYear} Manu Morante</small>
    </footer>
  )
}
