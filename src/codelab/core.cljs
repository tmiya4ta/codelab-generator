(ns codelab.core
  "CLI entry point for Codelab DOCX generator"
  (:require ["fs" :as fs]
            ["path" :as path]
            [codelab.meta :as meta]
            [codelab.parser :as parser]
            [codelab.docx :as docx]))

(def version "1.0.0")

(defn- print-usage []
  (println "
Codelab DOCX Generator v" version "

Usage: codelab-gen <input.md> [-o <output.docx>]

Arguments:
  input.md     Input markdown file (Codelab format)
  -o, --output Output DOCX file (default: <input>.docx)

Example:
  codelab-gen my-codelab.md -o my-codelab.docx

Markdown Format:
  ---
  id: my-codelab
  summary: Learn the basics
  categories: Web
  ---

  # Codelab Title

  ## Step 1
  Duration: 5:00

  Content here...
"))

(defn- parse-args
  "Parse command line arguments"
  [args]
  (loop [remaining args
         result {:input nil :output nil}]
    (if (empty? remaining)
      result
      (let [arg (first remaining)
            rest-args (rest remaining)]
        (cond
          (or (= arg "-h") (= arg "--help"))
          (assoc result :help true)

          (or (= arg "-o") (= arg "--output"))
          (recur (rest rest-args)
                 (assoc result :output (first rest-args)))

          (nil? (:input result))
          (recur rest-args
                 (assoc result :input arg))

          :else
          (recur rest-args result))))))

(defn- generate
  "Generate DOCX from markdown file"
  [input-path output-path]
  (if-not (fs/existsSync input-path)
    (do
      (println "Error: Input file not found:" input-path)
      (js/process.exit 1))
    (let [content (fs/readFileSync input-path "utf-8")
          base-dir (path/dirname (path/resolve input-path))
          {:keys [meta content]} (meta/parse-frontmatter content)
          ast (parser/parse content meta)]
      ;; generate-docx now returns a Promise
      (-> (docx/generate-docx ast base-dir)
          (.then (fn [doc] (docx/save-docx doc output-path)))
          (.catch (fn [err]
                    (js/console.error "Error:" err)
                    (js/process.exit 1)))))))

(defn main
  "Main entry point"
  [& args]
  (let [parsed (parse-args (js->clj (js/Array.from js/process.argv) :keywordize-keys false))
        ;; Skip first two args (node and script path)
        args-only (parse-args (drop 2 (js->clj (js/Array.from js/process.argv))))]
    (cond
      (:help args-only)
      (print-usage)

      (nil? (:input args-only))
      (do
        (println "Error: No input file specified")
        (print-usage)
        (js/process.exit 1))

      :else
      (let [input (:input args-only)
            output (or (:output args-only)
                       (str (path/basename input ".md") ".docx"))]
        (println "Converting:" input "→" output)
        (generate input output)))))
