(ns codelab.docx
  "DOCX generation using docx.js for Codelab format"
  (:require ["docx" :as docx]
            ["fs" :as fs]
            ["path" :as path]
            ["image-size" :refer [imageSize]]
            [clojure.string :as str]))

;; Sharp uses default export - require dynamically
(def sharp (js/require "sharp"))

;; Chart.js canvas renderer
(def ChartJSNodeCanvas (.-ChartJSNodeCanvas (js/require "chartjs-node-canvas")))

;; Import docx.js classes
(def Document (.-Document docx))
(def Paragraph (.-Paragraph docx))
(def TextRun (.-TextRun docx))
(def Table (.-Table docx))
(def TableRow (.-TableRow docx))
(def TableCell (.-TableCell docx))
(def ImageRun (.-ImageRun docx))
(def ExternalHyperlink (.-ExternalHyperlink docx))
(def HeadingLevel (.-HeadingLevel docx))
(def AlignmentType (.-AlignmentType docx))
(def WidthType (.-WidthType docx))
(def BorderStyle (.-BorderStyle docx))
(def ShadingType (.-ShadingType docx))
(def Packer (.-Packer docx))
(def LevelFormat (.-LevelFormat docx))
(def LevelSuffix (.-LevelSuffix docx))

;; Colors (from reference.docx)
(def color-gray "b7b7b7")
(def color-positive-bg "d9ead3")
(def color-negative-bg "fce5cd")
(def color-button-bg "6aa84f")
(def color-border "cccccc")
(def color-meta-key-bg "f3f3f3")

;; Fonts
(def font-default "Arial")
(def font-code "Courier New")
(def font-terminal "Consolas")
(def font-roboto "Roboto")
(def font-japanese "Calibri")
(def font-japanese-heading "Arial Unicode MS")
(def font-meta-japanese "SimSun")

;; Sizes (half-points: 20 = 10pt, 22 = 11pt, 42 = 21pt)
(def size-default 22)
(def size-table 20)
(def size-title 42)
(def size-heading1 40)
(def size-heading2 32)
(def size-heading3 28)

;; Heading colors
(def color-heading3 "434343")
(def color-heading4 "666666")

;; Line spacing (276 = 1.15, 240 = 1.0 line spacing in twips)
(def line-spacing-default 276)
(def line-spacing-single 240)

;; Numbering reference names
(def numbering-bullet-ref "codelab-bullet")
(def numbering-ordered-ref "codelab-ordered")

;; Body text size (Japanese text uses larger size)
(def size-body 24)

;; Code block border color (black)
(def color-code-border "000000")

;;; Numbering configuration (like reference DOCX)

(defn- make-numbering-config
  "Create numbering configuration for bullet and ordered lists"
  []
  {:config
   [;; Bullet list configuration
    {:reference numbering-bullet-ref
     :levels [{:level 0
               :format (.-BULLET LevelFormat)
               :text "●"
               :alignment (.-LEFT AlignmentType)
               :style {:paragraph {:indent {:left 720 :hanging 360}}
                       :run {:font {:name font-japanese}}}}
              {:level 1
               :format (.-BULLET LevelFormat)
               :text "○"
               :alignment (.-LEFT AlignmentType)
               :style {:paragraph {:indent {:left 1440 :hanging 360}}
                       :run {:font {:name font-japanese}}}}
              {:level 2
               :format (.-BULLET LevelFormat)
               :text "■"
               :alignment (.-LEFT AlignmentType)
               :style {:paragraph {:indent {:left 2160 :hanging 360}}
                       :run {:font {:name font-japanese}}}}]}
    ;; Ordered list configuration
    {:reference numbering-ordered-ref
     :levels [{:level 0
               :format (.-DECIMAL LevelFormat)
               :text "%1."
               :alignment (.-LEFT AlignmentType)
               :suffix (.-TAB LevelSuffix)
               :style {:paragraph {:indent {:left 720 :hanging 360}}
                       :run {:font {:name font-japanese}}}}
              {:level 1
               :format (.-LOWER_LETTER LevelFormat)
               :text "%2."
               :alignment (.-LEFT AlignmentType)
               :suffix (.-TAB LevelSuffix)
               :style {:paragraph {:indent {:left 1440 :hanging 360}}
                       :run {:font {:name font-japanese}}}}
              {:level 2
               :format (.-LOWER_ROMAN LevelFormat)
               :text "%3."
               :alignment (.-RIGHT AlignmentType)
               :suffix (.-TAB LevelSuffix)
               :style {:paragraph {:indent {:left 2160 :hanging 360}}
                       :run {:font {:name font-japanese}}}}]}]})

;;; Helper functions

(defn- make-text-run
  "Create a TextRun with options"
  [text opts]
  (TextRun. (clj->js (merge {:text text} opts))))

(defn- make-paragraph
  "Create a Paragraph with children and options"
  [children opts]
  (Paragraph. (clj->js (merge {:children children} opts))))

;;; Image processing with auto-trim

(defn- collect-image-paths
  "Recursively collect all image paths from AST"
  [node base-dir]
  (cond
    (map? node)
    (case (:type node)
      (:image :image-block)
      (let [src (:src node)
            full-path (if (path/isAbsolute src)
                        src
                        (path/join base-dir src))]
        (if (fs/existsSync full-path)
          [full-path]
          []))
      ;; Recurse into children
      (concat
       (collect-image-paths (:children node) base-dir)
       (collect-image-paths (:content node) base-dir)))

    (sequential? node)
    (mapcat #(collect-image-paths % base-dir) node)

    :else
    []))

(defn- process-image-async
  "Process a single image with sharp: trim whitespace and return buffer"
  [image-path]
  (-> (sharp image-path)
      (.trim)
      (.toBuffer)))

(defn process-all-images
  "Process all images in AST, returns Promise of {path -> buffer} map"
  [ast base-dir]
  (let [paths (distinct (collect-image-paths ast base-dir))]
    (if (empty? paths)
      (js/Promise.resolve {})
      (-> (js/Promise.all
           (clj->js (map (fn [p]
                           (-> (process-image-async p)
                               (.then (fn [buf] #js [p buf]))
                               (.catch (fn [_err]
                                         ;; On error, use original file
                                         #js [p (fs/readFileSync p)]))))
                         paths)))
          (.then (fn [results]
                   (into {} (map (fn [r] [(aget r 0) (aget r 1)]) results))))))))

;;; Inline element conversion

(defn- inline-node->runs
  "Convert inline AST node to TextRun(s)"
  [node base-dir & [processed-images]]
  (case (:type node)
    :text
    [(make-text-run (:content node)
                    (cond-> {:font {:name font-japanese}
                             :size size-body}
                      (:bold node) (assoc :bold true)
                      (:italic node) (assoc :italics true)))]

    :bold
    (mapcat #(inline-node->runs % base-dir processed-images)
            (map (fn [child]
                   (if (= (:type child) :text)
                     (assoc child :bold true)
                     child))
                 (:children node)))

    :italic
    (mapcat #(inline-node->runs % base-dir processed-images)
            (map (fn [child]
                   (if (= (:type child) :text)
                     (assoc child :italic true)
                     child))
                 (:children node)))

    :code
    [(make-text-run (:content node) {:font {:name font-code}})]

    :link
    [(ExternalHyperlink. (clj->js {:link (:href node)
                                   :children [(make-text-run (:text node)
                                                             {:style "Hyperlink"})]}))]

    :image
    (let [src (:src node)
          full-path (if (path/isAbsolute src)
                      src
                      (path/join base-dir src))]
      (if (fs/existsSync full-path)
        (let [;; Use pre-processed (trimmed) buffer if available, else read from file
              buffer (or (get processed-images full-path)
                         (fs/readFileSync full-path))
              dimensions (imageSize buffer)
              orig-width (.-width dimensions)
              orig-height (.-height dimensions)
              ;; Max width 500px, keep aspect ratio
              max-width 500
              scale (if (> orig-width max-width)
                      (/ max-width orig-width)
                      1)
              width (js/Math.round (* orig-width scale))
              height (js/Math.round (* orig-height scale))
              ext (str/lower-case (path/extname full-path))
              img-type (case ext
                         ".png" "png"
                         ".jpg" "jpg"
                         ".jpeg" "jpg"
                         ".gif" "gif"
                         "png")]
          [(ImageRun. (clj->js {:data buffer
                                :transformation {:width width :height height}
                                :type img-type}))])
        [(make-text-run (str "[Image not found: " src "]") {:color "FF0000"})]))

    ;; Default: treat as text
    [(make-text-run (str node) {})]))

(defn- inline-nodes->runs
  "Convert multiple inline nodes to TextRuns"
  [nodes base-dir & [processed-images]]
  (vec (mapcat #(inline-node->runs % base-dir processed-images) nodes)))

;;; Block element conversion

(declare block-node->elements)

(def ^:private table-border
  {:style (.-SINGLE BorderStyle)
   :size 8
   :color color-border
   :space 0})

(def ^:private table-borders
  {:top table-border
   :bottom table-border
   :left table-border
   :right table-border})

;; Black border for code blocks and infoboxes
(def ^:private code-border
  {:style (.-SINGLE BorderStyle)
   :size 8
   :color color-code-border
   :space 0})

(def ^:private code-borders
  {:top code-border
   :bottom code-border
   :left code-border
   :right code-border})

(def ^:private cell-margins
  {:top 100 :left 100 :bottom 100 :right 100})

(defn- meta-table-row
  "Create a row for the meta table"
  [key value]
  (TableRow. (clj->js {:children
                       [(TableCell. (clj->js {:children [(make-paragraph
                                                          [(make-text-run key {:bold true
                                                                               :size size-table})]
                                                          {:spacing {:line 240 :lineRule "auto"}})]
                                              :shading {:type (.-CLEAR ShadingType)
                                                        :fill color-meta-key-bg}
                                              :margins cell-margins
                                              :verticalAlign "top"
                                              :width {:size 2025 :type (.-DXA WidthType)}}))
                        (TableCell. (clj->js {:children [(make-paragraph
                                                          [(make-text-run (str value) {:font {:name font-roboto}
                                                                                       :size size-table})]
                                                          {:spacing {:line 240 :lineRule "auto"}})]
                                              :margins cell-margins
                                              :verticalAlign "top"
                                              :width {:size 7335 :type (.-DXA WidthType)}}))]})))

(def ^:private table-layout
  "Common table layout properties for Google Docs compatibility"
  {:layout (.-FIXED (.-TableLayoutType docx))
   :indent {:size 109 :type (.-DXA WidthType)}})

(defn- create-meta-table
  "Create the metadata table for Codelab"
  [meta]
  (let [key-map {"id" "URL"
                 "url" "URL"
                 "summary" "Summary"
                 "category" "Category"
                 "categories" "Category"
                 "environment" "Environment"
                 "environments" "Environment"
                 "status" "Status"
                 "feedback_link" "Feedback Link"
                 "analytics_account" "Analytics Account"
                 "author" "Author"
                 "authors" "Authors"
                 "tags" "Tags"}
        rows (for [[k v] meta
                   :when (get key-map k)]
               (meta-table-row (get key-map k) v))]
    (when (seq rows)
      (Table. (clj->js (merge table-layout
                              {:rows (vec rows)
                               :columnWidths [2025 7335]
                               :borders {:top table-border
                                         :bottom table-border
                                         :left table-border
                                         :right table-border
                                         :insideHorizontal table-border
                                         :insideVertical table-border}
                               :width {:size 9360 :type (.-DXA WidthType)}}))))))

(defn- create-single-cell-table
  "Create a single-cell table (for code blocks, infoboxes)"
  [children shading-color font-name]
  (Table.
   (clj->js
    (merge table-layout
           {:rows [(TableRow.
                    (clj->js
                     {:children
                      [(TableCell.
                        (clj->js
                         {:children children
                          :borders code-borders
                          :shading (when shading-color
                                     {:type (.-CLEAR ShadingType)
                                      :fill shading-color})
                          :margins {:top 100 :bottom 100 :left 100 :right 100}
                          :width {:size 9360 :type (.-DXA WidthType)}}))]}))]
            :columnWidths [9360]
            :borders {:top code-border
                      :bottom code-border
                      :left code-border
                      :right code-border
                      :insideH code-border
                      :insideV code-border}
            :width {:size 9360 :type (.-DXA WidthType)}}))))

(defn- create-code-block
  "Create a code block (single-cell table with Fira Mono)"
  [content lang]
  (let [lines (str/split-lines content)
        font (if (= lang "console") font-terminal font-code)
        paragraphs (map (fn [line]
                          (make-paragraph
                           [(make-text-run (if (str/blank? line) " " line)
                                           {:font {:name font}})]
                           {:spacing {:line line-spacing-single :lineRule "auto"}}))
                        lines)]
    (create-single-cell-table (vec paragraphs) nil font)))

(defn- create-infobox
  "Create an infobox (single-cell table with background color)"
  [children box-type base-dir processed-images]
  (let [bg-color (if (= box-type :infobox-positive)
                   color-positive-bg
                   color-negative-bg)
        content (mapcat #(block-node->elements % base-dir processed-images) children)]
    (create-single-cell-table (vec content) bg-color nil)))

(defn- create-list
  "Create a list (unordered or ordered) using proper DOCX numbering"
  [items ordered base-dir processed-images & [start-number instance-id]]
  (let [reference (if ordered numbering-ordered-ref numbering-bullet-ref)
        instance (or instance-id 0)]
    (map
     (fn [item]
       (let [runs (inline-nodes->runs (:children item) base-dir processed-images)
             level (or (:level item) 0)]
         (make-paragraph
          runs
          {:numbering {:reference reference
                       :level level
                       :instance instance}
           :spacing {:line line-spacing-single :lineRule "auto"}})))
     items)))

(defn- create-duration-para
  "Create Duration paragraph with gray text"
  [value]
  (make-paragraph
   [(make-text-run (str "Duration: " value) {:color color-gray})]
   {}))

(defn- create-environment-para
  "Create Environment paragraph with gray text"
  [value]
  (make-paragraph
   [(make-text-run (str "Environment: " value) {:color color-gray})]
   {}))

(defn- create-content-table
  "Create a simple content table from rows (for markdown tables)"
  [rows]
  (when (seq rows)
    (let [num-cols (count (first rows))
          col-width (js/Math.floor (/ 9360 num-cols))
          table-rows (map (fn [row]
                            (TableRow.
                             (clj->js
                              {:children
                               (map (fn [cell]
                                      (let [runs (if (string? cell)
                                                   [(make-text-run cell {:font {:name font-japanese}
                                                                         :size size-body})]
                                                   (inline-nodes->runs cell nil nil))]
                                        (TableCell.
                                         (clj->js
                                          {:children [(make-paragraph
                                                       runs
                                                       {:spacing {:line line-spacing-single :lineRule "auto"}})]
                                           :borders code-borders
                                           :margins cell-margins
                                           :width {:size col-width :type (.-DXA WidthType)}}))))
                                    row)})))
                          rows)]
      (Table.
       (clj->js
        (merge table-layout
               {:rows (vec table-rows)
                :columnWidths (vec (repeat num-cols col-width))
                :borders {:top code-border
                          :bottom code-border
                          :left code-border
                          :right code-border
                          :insideH code-border
                          :insideV code-border}
                :width {:size 9360 :type (.-DXA WidthType)}}))))))

;; Chart colors
(def chart-colors
  ["#4285F4" "#EA4335" "#FBBC04" "#34A853" "#FF6D01" "#46BDC6"
   "#7B1FA2" "#C2185B" "#00897B" "#6D4C41"])

(defn render-chart
  "Render a chart definition to a PNG buffer (returns Promise)"
  [chart-json]
  (let [config (js/JSON.parse chart-json)
        chart-type (or (.-type config) "bar")
        title (.-title config)
        labels (js->clj (.-labels config))
        datasets (js->clj (.-datasets config) :keywordize-keys true)
        width (or (.-width config) 700)
        height (or (.-height config) 400)
        canvas (ChartJSNodeCanvas. (clj->js {:width width :height height
                                              :backgroundColour "white"}))
        chart-config (clj->js
                      {:type chart-type
                       :data {:labels labels
                              :datasets (vec (map-indexed
                                             (fn [i ds]
                                               {:label (:label ds)
                                                :data (:data ds)
                                                :backgroundColor (nth chart-colors (mod i (count chart-colors)))
                                                :borderColor (nth chart-colors (mod i (count chart-colors)))
                                                :borderWidth 1})
                                             datasets))}
                       :options {:responsive false
                                 :plugins {:title {:display (some? title)
                                                   :text (or title "")
                                                   :font {:size 16}}
                                           :legend {:position "bottom"}}
                                 :scales {:y {:beginAtZero true}}}})]
    (.renderToBuffer canvas chart-config)))

(defn- create-image-block
  "Create an image block with empty lines before and after"
  [node base-dir processed-images]
  (let [src (:src node)
        full-path (if (path/isAbsolute src)
                    src
                    (path/join base-dir src))]
    (if (fs/existsSync full-path)
      (let [buffer (or (get processed-images full-path)
                       (fs/readFileSync full-path))
            dimensions (imageSize buffer)
            orig-width (.-width dimensions)
            orig-height (.-height dimensions)
            max-width 500
            scale (if (> orig-width max-width)
                    (/ max-width orig-width)
                    1)
            width (js/Math.round (* orig-width scale))
            height (js/Math.round (* orig-height scale))
            ext (str/lower-case (path/extname full-path))
            img-type (case ext
                       ".png" "png"
                       ".jpg" "jpg"
                       ".jpeg" "jpg"
                       ".gif" "gif"
                       "png")]
        [(make-paragraph [] {})  ;; Empty line before
         (make-paragraph
          [(ImageRun. (clj->js {:data buffer
                                :transformation {:width width :height height}
                                :type img-type}))]
          {})
         (make-paragraph [] {})])  ;; Empty line after
      [(make-paragraph
        [(make-text-run (str "[Image not found: " src "]") {:color "FF0000"})]
        {})])))

(defn block-node->elements
  "Convert block AST node to DOCX elements"
  [node base-dir & [processed-images]]
  (case (:type node)
    :title
    [(make-paragraph
      [(make-text-run (:content node) {:size size-title
                                       :font {:name font-japanese-heading}})]
      {:heading (.-TITLE HeadingLevel)
       :spacing {:after 0 :line line-spacing-default :lineRule "auto"}})]

    :step
    (let [title-para (make-paragraph
                      [(make-text-run (:title node) {:size size-heading1
                                                     :font {:name font-japanese-heading}})]
                      {:heading (.-HEADING_1 HeadingLevel)
                       :spacing {:before 400 :after 120 :line line-spacing-default :lineRule "auto"}})
          children (mapcat #(block-node->elements % base-dir processed-images) (:children node))]
      (into [title-para] children))

    :heading
    (let [[level text-opts para-opts] (case (:level node)
                                        1 [(.-HEADING_1 HeadingLevel)
                                           {:size size-heading1
                                            :font {:name font-japanese-heading}}
                                           {:spacing {:before 400 :after 120}}]
                                        2 [(.-HEADING_2 HeadingLevel)
                                           {:size size-heading2
                                            :font {:name font-japanese-heading}}
                                           {:spacing {:before 360 :after 120}}]
                                        3 [(.-HEADING_3 HeadingLevel)
                                           {:size size-heading3 :color color-heading3
                                            :font {:name font-japanese-heading}}
                                           {:spacing {:before 320 :after 80}}]
                                        [(.-HEADING_4 HeadingLevel)
                                         {:color color-heading4
                                          :font {:name font-japanese-heading}}
                                         {:spacing {:before 280 :after 80}}])]
      [(make-paragraph
        [(make-text-run (:content node) text-opts)]
        (merge {:heading level} para-opts))])

    :paragraph
    [(make-paragraph (inline-nodes->runs (:children node) base-dir processed-images)
                     {:spacing {:line line-spacing-single :lineRule "auto"}})]

    :meta-duration
    [(create-duration-para (:value node))
     (make-paragraph [] {})]  ;; Empty line after Duration

    :meta-environment
    [(create-environment-para (:value node))
     (make-paragraph [] {})]  ;; Empty line after Environment

    :code-block
    [(make-paragraph [] {})  ;; Empty line before
     (create-code-block (:content node) (:lang node))
     (make-paragraph [] {})]  ;; Empty line after

    :terminal
    [(make-paragraph [] {})  ;; Empty line before
     (create-code-block (:content node) "console")
     (make-paragraph [] {})]  ;; Empty line after

    :infobox-positive
    [(make-paragraph [] {})  ;; Empty line before
     (create-infobox (:children node) :infobox-positive base-dir processed-images)
     (make-paragraph [] {})]  ;; Empty line after

    :infobox-negative
    [(make-paragraph [] {})  ;; Empty line before
     (create-infobox (:children node) :infobox-negative base-dir processed-images)
     (make-paragraph [] {})]  ;; Empty line after

    :list
    (create-list (:children node) (:ordered node) base-dir processed-images nil (:instance-id node))

    :table
    [(make-paragraph [] {})  ;; Empty line before
     (create-content-table (:rows node))
     (make-paragraph [] {})]  ;; Empty line after

    :chart
    (let [chart-key (str "__chart_" (:chart-index node))
          buf (get processed-images chart-key)]
      (if buf
        (let [dimensions (imageSize buf)
              width (min 500 (.-width dimensions))
              scale (/ width (.-width dimensions))
              height (js/Math.round (* (.-height dimensions) scale))]
          [(make-paragraph [] {})
           (make-paragraph
            [(ImageRun.
              (clj->js {:data buf
                        :transformation {:width width :height height}
                        :type "png"}))]
            {:alignment (.-CENTER AlignmentType)})
           (make-paragraph [] {})])
        [(make-paragraph
          [(make-text-run "[Chart render failed]" {:color "FF0000"})]
          {})]))

    :image-block
    (create-image-block node base-dir processed-images)

    ;; Default: empty
    []))

;;; Document generation

(defn- annotate-list-instances
  "Annotate lists with instance IDs for proper numbering.
   Instance ID changes when a new step or heading begins (restarts numbering).
   Lists within the same section share an instance ID (continues numbering)."
  ([nodes] (annotate-list-instances nodes 0))
  ([nodes start-instance]
   (loop [remaining nodes
          result []
          instance start-instance]
     (if (empty? remaining)
       [result instance]
       (let [node (first remaining)
             rest-nodes (rest remaining)]
         (cond
           ;; List: add instance-id
           (= (:type node) :list)
           (let [annotated (assoc node :instance-id instance)]
             (recur rest-nodes
                    (conj result annotated)
                    instance))

           ;; Step: recursively process children with new instance
           (= (:type node) :step)
           (let [[processed-children next-instance] (annotate-list-instances (:children node) (inc instance))]
             (recur rest-nodes
                    (conj result (assoc node :children processed-children))
                    next-instance))

           ;; Heading (###, ####, etc.): increment instance for next list
           (= (:type node) :heading)
           (recur rest-nodes
                  (conj result node)
                  (inc instance))

           ;; Other nodes: pass through, keep instance
           :else
           (recur rest-nodes
                  (conj result node)
                  instance)))))))

(defn- collect-charts*
  "Recursively collect chart nodes from AST children"
  [nodes]
  (reduce (fn [acc node]
            (if (= :chart (:type node))
              (conj acc (:content node))
              (if (:children node)
                (into acc (collect-charts* (:children node)))
                acc)))
          []
          nodes))

(defn- collect-charts
  "Collect all chart nodes from AST, returns list of {:index N :content json}"
  [ast]
  (let [contents (collect-charts* (:children ast))]
    (map-indexed (fn [i c] {:index i :content c}) contents)))

(defn- process-all-charts
  "Render all charts to PNG buffers. Returns Promise of {index -> buffer}"
  [ast]
  (let [charts (collect-charts ast)]
    (if (empty? charts)
      (js/Promise.resolve {})
      (-> (js/Promise.all
           (clj->js (map (fn [{:keys [index content]}]
                           (-> (render-chart content)
                               (.then (fn [buf] #js [index buf]))
                               (.catch (fn [err]
                                         (js/console.error "Chart render error:" err)
                                         #js [index nil]))))
                         charts)))
          (.then (fn [results]
                   (into {} (map (fn [r] [(aget r 0) (aget r 1)]) results))))))))

(defn- annotate-chart-indices
  "Add :chart-index to each :chart node in AST (depth-first)"
  [nodes]
  (let [counter (atom 0)]
    (letfn [(walk [nodes]
              (mapv (fn [node]
                      (if (= :chart (:type node))
                        (let [idx @counter]
                          (swap! counter inc)
                          (assoc node :chart-index idx))
                        (if (:children node)
                          (update node :children walk)
                          node)))
                    nodes))]
      (walk nodes))))

(defn generate-docx
  "Generate DOCX document from AST. Returns a Promise."
  [ast base-dir]
  (-> (js/Promise.all #js [(process-all-images ast base-dir)
                            (process-all-charts ast)])
      (.then (fn [results]
               (let [processed-images (aget results 0)
                     chart-buffers (aget results 1)
                     ;; Merge chart buffers into processed-images as "__chart_N"
                     all-images (reduce (fn [m [idx buf]]
                                          (if buf
                                            (assoc m (str "__chart_" idx) buf)
                                            m))
                                        processed-images
                                        chart-buffers)
                     ;; Annotate chart nodes with indices
                     annotated-children (annotate-chart-indices (:children ast))
                     [children _] (annotate-list-instances annotated-children)
                     content-elements (mapcat #(block-node->elements % base-dir all-images) children)]
                 (Document.
                  (clj->js
                   {:styles {:default {:document {:run {:font {:name font-default}
                                                        :size size-default}
                                                  :paragraph {:spacing {:line line-spacing-default
                                                                        :lineRule "auto"}}}}}
                    :numbering (make-numbering-config)
                    :sections [{:children (vec content-elements)}]})))))))

(defn save-docx
  "Save DOCX document to file"
  [doc output-path]
  (-> (.toBuffer Packer doc)
      (.then (fn [buffer]
               (fs/writeFileSync output-path buffer)
               (println "Generated:" output-path)))
      (.catch (fn [err]
                (js/console.error "Error generating DOCX:" err)))))
