(ns codelab.html
  "Render the Codelab AST to claat-compatible HTML (google-codelab elements)"
  (:require ["fs" :as fs]
            ["path" :as path]
            ["crypto" :as crypto]
            [clojure.string :as str]
            [codelab.docx :as docx]))

;;; Resources (template and codelab-elements), shipped next to dist/

(defn- resource-dir []
  (path/join js/__dirname ".." "resources" "html"))

(def ^:private element-files
  ["codelab-elements.css" "codelab-elements.js"
   "custom-elements.min.js" "native-shim.js" "prettify.js"])

;;; Escaping

(defn- escape [s]
  (-> (str s)
      (str/replace "&" "&amp;")
      (str/replace "<" "&lt;")
      (str/replace ">" "&gt;")
      (str/replace "\"" "&quot;")))

;;; Images: copy into img/ under a content hash name (like claat)

(defn- image-name [buf ext]
  (str (subs (-> (crypto/createHash "sha1") (.update buf) (.digest "hex")) 0 16) ext))

(defn- copy-image!
  "Copy an image into out-dir/img/. Returns the relative src, or nil if missing."
  [src base-dir out-dir]
  (let [full (if (path/isAbsolute src) src (path/join base-dir src))]
    (if (fs/existsSync full)
      (let [buf (fs/readFileSync full)
            name (image-name buf (str/lower-case (path/extname full)))]
        (fs/writeFileSync (path/join out-dir "img" name) buf)
        (str "img/" name))
      (do (js/console.warn "Warning: image not found:" full)
          nil))))

(defn- write-buffer-image! [buf out-dir]
  (let [name (image-name buf ".png")]
    (fs/writeFileSync (path/join out-dir "img" name) buf)
    (str "img/" name)))

;;; Inline nodes

(declare inline->html)

(defn- img-tag [alt src]
  (str "<img alt=\"" (escape alt) "\" src=\"" (escape src) "\">"))

(defn- inline-node->html [node ctx]
  (case (:type node)
    :text (escape (:content node))
    :code (str "<code>" (escape (:content node)) "</code>")
    :bold (str "<strong>" (inline->html (:children node) ctx) "</strong>")
    :italic (str "<em>" (inline->html (:children node) ctx) "</em>")
    :link (str "<a href=\"" (escape (:href node)) "\" target=\"_blank\">"
               (escape (:text node)) "</a>")
    :image (if-let [src (copy-image! (:src node) (:base-dir ctx) (:out-dir ctx))]
             (img-tag (:alt node) src)
             "")
    (escape (:content node))))

(defn- inline->html [nodes ctx]
  (apply str (map #(inline-node->html % ctx) nodes)))

(defn- cell->html [cell ctx]
  (if (string? cell) (escape cell) (inline->html cell ctx)))

;;; Block nodes

(defn- duration-minutes
  "Duration: mm:ss or hh:mm:ss → whole minutes"
  [value]
  (let [parts (map #(js/parseInt % 10) (str/split value #":"))
        [h m s] (case (count parts)
                  3 parts
                  2 (cons 0 parts)
                  [0 (first parts) 0])]
    (js/Math.round (+ (* 60 h) m (/ s 60)))))

(declare blocks->html)

(defn- block->html
  "Render one block. `ctx` carries :base-dir, :out-dir, :charts and an
   atom :ol-count that continues ordered-list numbering within a section
   (the same rule as the DOCX output)."
  [node ctx]
  (case (:type node)
    :paragraph (str "<p>" (inline->html (:children node) ctx) "</p>\n")

    :heading (do (reset! (:ol-count ctx) 0)
                 (let [level (min 6 (max 3 (:level node)))]
                   (str "<h" level ">" (inline->html (:children node) ctx) "</h" level ">\n")))

    :list (let [items (apply str (map #(str "<li>" (inline->html (:children %) ctx) "</li>\n")
                                      (:children node)))]
            (if (:ordered node)
              (let [start (inc @(:ol-count ctx))]
                (swap! (:ol-count ctx) + (count (:children node)))
                (str "<ol" (when (> start 1) (str " start=\"" start "\"")) ">\n" items "</ol>\n"))
              (str "<ul>\n" items "</ul>\n")))

    (:code-block :terminal)
    (let [lang (:lang node)]
      (str "<pre><code"
           (when (seq lang) (str " language=\"" (escape lang) "\" class=\"" (escape lang) "\""))
           ">" (escape (:content node)) "</code></pre>\n"))

    :table (str "<table>\n"
                (apply str (map (fn [row]
                                  (str "<tr>"
                                       (apply str (map #(str "<td>" (cell->html % ctx) "</td>") row))
                                       "</tr>\n"))
                                (:rows node)))
                "</table>\n")

    :image-block (if-let [src (copy-image! (:src node) (:base-dir ctx) (:out-dir ctx))]
                   (str "<p class=\"image-container\">" (img-tag (:alt node) src) "</p>\n")
                   "")

    :chart (if-let [buf (get (:charts ctx) (:chart-index node))]
             (str "<p class=\"image-container\">"
                  (img-tag "" (write-buffer-image! buf (:out-dir ctx))) "</p>\n")
             "")

    :infobox-positive (str "<aside class=\"special\">" (blocks->html (:children node) ctx) "</aside>\n")
    :infobox-negative (str "<aside class=\"warning\">" (blocks->html (:children node) ctx) "</aside>\n")

    ;; Duration / Environment are attributes of the step, not content
    (:meta-duration :meta-environment) ""

    ""))

(defn- blocks->html [nodes ctx]
  (apply str (map #(block->html % ctx) nodes)))

(defn- step->html [step ctx]
  (reset! (:ol-count ctx) 0)
  (let [duration (some #(when (= :meta-duration (:type %)) (:value %)) (:children step))]
    (str "<google-codelab-step label=\"" (escape (:title step)) "\" duration=\""
         (if duration (duration-minutes duration) 0) "\">\n"
         (blocks->html (:children step) ctx)
         "</google-codelab-step>\n")))

;;; Charts (rendered with the same function as the DOCX output)

(defn- annotate-chart-indices [nodes counter]
  (mapv (fn [node]
          (cond
            (= :chart (:type node)) (let [i @counter] (swap! counter inc) (assoc node :chart-index i))
            (:children node) (update node :children annotate-chart-indices counter)
            :else node))
        nodes))

(defn- collect-charts [nodes]
  (mapcat (fn [node]
            (cond
              (= :chart (:type node)) [node]
              (:children node) (collect-charts (:children node))
              :else []))
          nodes))

(defn- render-charts [charts]
  (-> (js/Promise.all
       (clj->js (map (fn [c]
                       (-> (docx/render-chart (:content c))
                           (.then (fn [buf] #js [(:chart-index c) buf]))
                           (.catch (fn [err]
                                     (js/console.error "Chart render error:" err)
                                     #js [(:chart-index c) nil]))))
                     charts)))
      (.then (fn [results] (into {} (map (fn [r] [(aget r 0) (aget r 1)]) results))))))

;;; Entry point

(defn generate-html
  "Write out-dir/index.html, out-dir/img/ and out-dir/elements/. Returns a Promise."
  [ast base-dir out-dir]
  (let [children (annotate-chart-indices (:children ast) (atom 0))]
    (-> (render-charts (collect-charts children))
        (.then
         (fn [charts]
           (fs/mkdirSync (path/join out-dir "img") #js {:recursive true})
           (fs/mkdirSync (path/join out-dir "elements") #js {:recursive true})
           (doseq [f element-files]
             (fs/copyFileSync (path/join (resource-dir) "elements" f)
                              (path/join out-dir "elements" f)))
           (let [meta (:meta ast)
                 ctx {:base-dir base-dir :out-dir out-dir :charts charts :ol-count (atom 0)}
                 title (or (some #(when (= :title (:type %)) (:content %)) children) "")
                 steps (apply str (map #(step->html % ctx) (filter #(= :step (:type %)) children)))
                 template (fs/readFileSync (path/join (resource-dir) "template.html") "utf-8")
                 ;; Replace with functions so that `$&` etc. in the content stay literal
                 html (-> template
                          (str/replace "{{title}}" (constantly (escape title)))
                          (str/replace "{{id}}" (constantly (escape (get meta "id" ""))))
                          (str/replace "{{feedback}}" (constantly (escape (or (get meta "feedback_link")
                                                                              (get meta "feedback") ""))))
                          (str/replace "{{steps}}" (constantly steps)))
                 index (path/join out-dir "index.html")]
             (fs/writeFileSync index html)
             (println "Generated:" index)))))))
