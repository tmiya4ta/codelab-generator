(ns codelab.core
  "CLI entry point for Codelab DOCX / HTML generator"
  (:require ["fs" :as fs]
            ["path" :as path]
            [codelab.meta :as meta]
            [codelab.parser :as parser]
            [codelab.docx :as docx]
            [codelab.html :as html]))

(def version "1.1.0")

(defn- print-usage []
  (println "
Codelab Generator v" version "

Usage: codelab-gen <input.md> [-f docx|html] [-o <output>]

Arguments:
  input.md     Input markdown file (Codelab format)
  -f, --format Output format: docx (default) or html
  -o, --output DOCX: output file (default: <input>.docx)
               HTML: output directory (default: <id>/), gets index.html, img/ and elements/

Example:
  codelab-gen my-codelab.md -o my-codelab.docx
  codelab-gen my-codelab.md -f html -o out/my-codelab

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
         result {:input nil :output nil :format "docx"}]
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

          (or (= arg "-f") (= arg "--format"))
          (recur (rest rest-args)
                 (assoc result :format (first rest-args)))

          (nil? (:input result))
          (recur rest-args
                 (assoc result :input arg))

          :else
          (recur rest-args result))))))

(defn- generate
  "Generate DOCX or HTML from markdown file"
  [input-path output-path fmt]
  (if-not (fs/existsSync input-path)
    (do
      (println "Error: Input file not found:" input-path)
      (js/process.exit 1))
    (let [content (fs/readFileSync input-path "utf-8")
          base-dir (path/dirname (path/resolve input-path))
          {:keys [meta content]} (meta/parse-frontmatter content)
          ast (parser/parse content meta)]
      (-> (if (= fmt "html")
            (html/generate-html ast base-dir (or output-path (get meta "id")
                                                 (path/basename input-path ".md")))
            ;; generate-docx returns a Promise
            (-> (docx/generate-docx ast base-dir)
                (.then (fn [doc] (docx/save-docx doc (or output-path
                                                         (str (path/basename input-path ".md") ".docx")))))))
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

      (not (#{"docx" "html"} (:format args-only)))
      (do
        (println "Error: Unknown format:" (:format args-only))
        (print-usage)
        (js/process.exit 1))

      (nil? (:input args-only))
      (do
        (println "Error: No input file specified")
        (print-usage)
        (js/process.exit 1))

      :else
      (let [{:keys [input output format]} args-only]
        (println "Converting:" input "→" (or output format))
        (generate input output format)))))
