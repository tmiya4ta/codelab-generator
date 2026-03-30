(ns codelab.parser
  "Markdown parser for Codelab format"
  (:require [clojure.string :as str]))

;; Node types for AST
(def node-types
  #{:document :step :paragraph :heading
    :text :bold :italic :code :code-block :terminal :chart
    :list :list-item :link :image :image-block :table
    :infobox-positive :infobox-negative
    :meta-duration :meta-environment})

;;; Inline parsing

(declare parse-inline)

(defn- escape-regex
  "Escape special regex characters"
  [s]
  (str/replace s #"[.*+?^${}()|\[\]\\]" "\\$&"))

(def ^:private inline-patterns
  "Patterns for inline elements (order matters)"
  [;; Code (backticks) - must be before bold/italic
   {:pattern #"`([^`]+)`"
    :type :code
    :extract (fn [[_ content]] {:type :code :content content})}
   ;; Bold
   {:pattern #"\*\*([^*]+)\*\*"
    :type :bold
    :extract (fn [[_ content]] {:type :bold :children (parse-inline content)})}
   ;; Italic
   {:pattern #"\*([^*]+)\*"
    :type :italic
    :extract (fn [[_ content]] {:type :italic :children (parse-inline content)})}
   ;; Image
   {:pattern #"!\[([^\]]*)\]\(([^)]+)\)"
    :type :image
    :extract (fn [[_ alt src]] {:type :image :alt alt :src src})}
   ;; Link
   {:pattern #"\[([^\]]+)\]\(([^)]+)\)"
    :type :link
    :extract (fn [[_ text href]] {:type :link :text text :href href})}])

(defn parse-inline
  "Parse inline elements from text"
  [text]
  (if (str/blank? text)
    []
    (loop [remaining text
           result []]
      (if (str/blank? remaining)
        result
        (let [matches (for [{:keys [pattern extract]} inline-patterns
                            :let [m (re-find pattern remaining)]
                            :when m]
                        {:match m
                         :index (str/index-of remaining (first m))
                         :extract extract})
              first-match (first (sort-by :index matches))]
          (if first-match
            (let [{:keys [match index extract]} first-match
                  before (subs remaining 0 index)
                  after (subs remaining (+ index (count (first match))))
                  node (extract match)]
              (recur after
                     (cond-> result
                       (seq before) (conj {:type :text :content before})
                       true (conj node))))
            (conj result {:type :text :content remaining})))))))

;;; Block parsing

(defn- parse-heading
  "Parse heading line"
  [line]
  (when-let [[_ hashes content] (re-matches #"^(#{1,6})\s+(.+)$" line)]
    {:type :heading
     :level (count hashes)
     :content (str/trim content)
     :children (parse-inline (str/trim content))}))

(defn- parse-list-item
  "Parse list item line"
  [line]
  (cond
    (re-matches #"^\s*[-*]\s+.+" line)
    {:type :list-item
     :ordered false
     :children (parse-inline (str/replace line #"^\s*[-*]\s+" ""))}

    (re-matches #"^\s*\d+\.\s+.+" line)
    {:type :list-item
     :ordered true
     :children (parse-inline (str/replace line #"^\s*\d+\.\s+" ""))}

    :else nil))

(defn- parse-duration
  "Parse Duration: mm:ss or Duration: hh:mm:ss"
  [line]
  (when-let [[_ duration] (re-matches #"^Duration:\s*(.+)$" line)]
    {:type :meta-duration :value (str/trim duration)}))

(defn- parse-environment
  "Parse Environment: env1, env2"
  [line]
  (when-let [[_ env] (re-matches #"^Environment:\s*(.+)$" line)]
    {:type :meta-environment :value (str/trim env)}))

(defn- starts-code-block?
  "Check if line starts a code block"
  [line]
  (re-matches #"^```(\w*).*$" line))

(defn- ends-code-block?
  "Check if line ends a code block"
  [line]
  (= line "```"))

(defn- parse-infobox-start
  "Check if line starts an infobox (> [!NOTE] or > [!WARNING])"
  [line]
  (cond
    (re-matches #"^>\s*\[!NOTE\].*$" line) :infobox-positive
    (re-matches #"^>\s*\[!WARNING\].*$" line) :infobox-negative
    :else nil))

(defn- parse-blockquote-line
  "Parse a blockquote line, removing the > prefix"
  [line]
  (if (str/starts-with? line ">")
    (str/replace line #"^>\s?" "")
    nil))

(defn- table-line?
  "Check if line is a table row (starts and ends with |)"
  [line]
  (and (str/starts-with? (str/trim line) "|")
       (str/ends-with? (str/trim line) "|")))

(defn- table-separator?
  "Check if line is a table separator (|---|---|)"
  [line]
  (re-matches #"^\|[\s\-:|]+\|$" (str/trim line)))

(defn- parse-table-row
  "Parse a table row into cells (with inline markup)"
  [line]
  (let [trimmed (str/trim line)
        ;; Remove leading and trailing |
        inner (subs trimmed 1 (dec (count trimmed)))
        cells (str/split inner #"\|")]
    (mapv (fn [cell]
            (let [text (str/trim cell)
                  inlines (parse-inline text)]
              (if (and (= 1 (count inlines))
                       (= :text (:type (first inlines))))
                text  ;; plain text - keep as string for backwards compat
                inlines)))  ;; has inline markup - return parsed nodes
          cells)))

(defn- image-only-paragraph?
  "Check if a line contains only an image"
  [line]
  (when-let [[_ alt src] (re-matches #"^\s*!\[([^\]]*)\]\(([^)]+)\)\s*$" line)]
    {:alt alt :src src}))

;;; Main parser

(defn- collect-code-block
  "Collect lines until end of code block"
  [lines lang]
  (loop [remaining lines
         code-lines []]
    (if (empty? remaining)
      {:code (str/join "\n" code-lines)
       :remaining []
       :lang lang}
      (let [line (first remaining)]
        (if (ends-code-block? line)
          {:code (str/join "\n" code-lines)
           :remaining (rest remaining)
           :lang lang}
          (recur (rest remaining) (conj code-lines line)))))))

(defn- collect-infobox
  "Collect lines until end of infobox (non-blockquote line)"
  [lines box-type]
  (loop [remaining lines
         content-lines []]
    (if (empty? remaining)
      {:content (str/join "\n" content-lines)
       :remaining []
       :type box-type}
      (let [line (first remaining)
            content (parse-blockquote-line line)]
        (if content
          (recur (rest remaining) (conj content-lines content))
          {:content (str/join "\n" content-lines)
           :remaining remaining
           :type box-type})))))

(defn- collect-list
  "Collect consecutive list items"
  [lines first-item]
  (loop [remaining lines
         items [first-item]]
    (if (empty? remaining)
      {:items items :remaining [] :ordered (:ordered first-item)}
      (let [line (first remaining)
            item (parse-list-item line)]
        (if (and item (= (:ordered item) (:ordered first-item)))
          (recur (rest remaining) (conj items item))
          {:items items :remaining remaining :ordered (:ordered first-item)})))))

(defn- collect-table
  "Collect consecutive table lines"
  [lines first-line]
  (loop [remaining lines
         rows [first-line]]
    (if (empty? remaining)
      {:rows rows :remaining []}
      (let [line (first remaining)]
        (if (table-line? line)
          (recur (rest remaining) (conj rows line))
          {:rows rows :remaining remaining})))))

(defn parse-blocks
  "Parse markdown content into block-level AST"
  [content]
  (let [lines (str/split-lines content)]
    (loop [remaining lines
           blocks []]
      (if (empty? remaining)
        blocks
        (let [line (first remaining)
              rest-lines (rest remaining)]
          (cond
            ;; Empty line
            (str/blank? line)
            (recur rest-lines blocks)

            ;; Heading
            (parse-heading line)
            (recur rest-lines (conj blocks (parse-heading line)))

            ;; Duration
            (parse-duration line)
            (recur rest-lines (conj blocks (parse-duration line)))

            ;; Environment
            (parse-environment line)
            (recur rest-lines (conj blocks (parse-environment line)))

            ;; Code block
            (starts-code-block? line)
            (let [[_ lang] (re-matches #"^```(\w*).*$" line)
                  {:keys [code remaining lang]} (collect-code-block rest-lines lang)]
              (if (= lang "chart")
                (recur remaining
                       (conj blocks {:type :chart
                                     :content code}))
                (let [block-type (if (= lang "console") :terminal :code-block)]
                  (recur remaining
                         (conj blocks {:type block-type
                                       :lang lang
                                       :content code})))))

            ;; Infobox
            (parse-infobox-start line)
            (let [box-type (parse-infobox-start line)
                  {:keys [content remaining type]} (collect-infobox rest-lines box-type)]
              (recur remaining
                     (conj blocks {:type type
                                   :children (parse-blocks content)})))

            ;; List item
            (parse-list-item line)
            (let [first-item (parse-list-item line)
                  {:keys [items remaining ordered]} (collect-list rest-lines first-item)]
              (recur remaining
                     (conj blocks {:type :list
                                   :ordered ordered
                                   :children items})))

            ;; Table
            (table-line? line)
            (let [{:keys [rows remaining]} (collect-table rest-lines line)
                  ;; Parse rows, skip separator row
                  parsed-rows (->> rows
                                   (remove table-separator?)
                                   (mapv parse-table-row))]
              (recur remaining
                     (conj blocks {:type :table
                                   :rows parsed-rows})))

            ;; Standalone image (with empty lines before/after)
            (image-only-paragraph? line)
            (let [{:keys [alt src]} (image-only-paragraph? line)]
              (recur rest-lines
                     (conj blocks {:type :image-block
                                   :alt alt
                                   :src src})))

            ;; Regular paragraph
            :else
            (recur rest-lines
                   (conj blocks {:type :paragraph
                                 :children (parse-inline line)}))))))))

(defn parse-steps
  "Parse content into Codelab steps (split by ## headings)"
  [content]
  (let [blocks (parse-blocks content)]
    (loop [remaining blocks
           steps []
           current-step nil]
      (if (empty? remaining)
        (if current-step
          (conj steps current-step)
          steps)
        (let [block (first remaining)
              rest-blocks (rest remaining)]
          (if (and (= (:type block) :heading)
                   (= (:level block) 2))
            ;; New step (## heading)
            (recur rest-blocks
                   (if current-step (conj steps current-step) steps)
                   {:type :step
                    :title (:content block)
                    :children []})
            ;; Content within step
            (if current-step
              (recur rest-blocks
                     steps
                     (update current-step :children conj block))
              ;; Content before first step (title or intro)
              (if (and (= (:type block) :heading)
                       (= (:level block) 1))
                (recur rest-blocks
                       (conj steps {:type :title :content (:content block)})
                       nil)
                (recur rest-blocks
                       (conj steps block)
                       nil)))))))))

(defn parse
  "Parse markdown content into Codelab AST"
  [content meta]
  {:type :document
   :meta meta
   :children (parse-steps content)})
