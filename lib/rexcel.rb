
module Rexcel
  class Worksheet
    attr_accessor :array, :name, :col_width, :col_default_width, :default_style
    DefaultStyle = {:colspan => nil, :rowspan => nil, :horizontal => "left", :vertical => "center", :bold => false, :italic => false, :underline => "none", :size => 12, :font => "Arial", :back_color => "none", :color => "#000000", :border => {:left => false, :right => false, :top => false, :bottom => false, :color => "#000000", :style => "continuous", :weight => 1}}
    def initialize(name)
      @array = Array.new
      @name = work_on_name(name)
      @col_width = Array.new
      @col_default_width = 60
      @default_style = {}
    end
    
    def work_on_name(name)
      name = name.gsub("\n", " ")
      invalid_char = [':', "\\", "/"]
      invalid_char.each do |char|
        name = name.gsub(char, "")
      end
      name
    end
    
    def add_line(element, params = {}, sort_list = nil, title = {:title => false, :params => {}})
      #if we have passed a active-record object, array is a array of object. We need the attributes 
      #of that object:
      if element.respond_to?(:attributes)
        line = element.attributes
      elsif (element.class == Array) || (element.class == Hash)
        line = element
      end
      #if the line is a hash and if we ask for title, we add a line with the titles. (keys of the hash)
      if title[:title]
        #let's get the titles in a array:
        title_array = ((sort_list.class == Array) &&(!sort_list.empty?)) ? sort_list : ((line.class == Hash) ? line.keys : [])
        if title[:params].class == Hash
          add_clone_line(title_array, title[:params])
        else
          @array << title_array
        end
      end
      #if the line is a hash we can sort it:
      if (line.class == Hash) && ((sort_list.class == Array) &&(!sort_list.empty?))
         sorted_line = sort_hash(line, sort_list)
      else
        sorted_line = line
      end
      if params.class == Hash
        add_clone_line(sorted_line, params)
      else
        @array << sorted_line
      end
    end
    
    def add_lines(array, params = {}, sort_list = nil, title = {:title => false, :params => {}})
      #if the line is a hash and if we ask for title, we add a line with the titles. (keys of the hash)
      if title[:title]
        #let's get the titles in a array:
        ok = false
        if ((sort_list.class == Array) &&(!sort_list.empty?))
          line = sort_list
          ok = true
        else
          if array[0].respond_to?(:attributes)
            title_array = array[0].attributes.keys
            line = title_array
            ok = true
          elsif array[0].class == Hash
            title_array = array[0].keys
            line = title_array
            ok = true
          end
        end
        if ok
          if title[:params].class == Hash
            add_clone_line(line, title[:params])
          else
            @array << line
          end
        end
      end
      
      for element in array
        #if we have passed a active-record object, array is a array of object. We need the attributes 
        #of that object:
        if element.respond_to?(:attributes)
          line = element.attributes
        elsif (element.class == Array) || (element.class == Hash)
          line = element
        end
        #if the line is a hash we can sort it:
        if (line.class == Hash) && ((sort_list.class == Array) &&(!sort_list.empty?))
           sorted_line = sort_hash(line, sort_list)
        else
          sorted_line = line
        end
        if params.class == Hash
          add_clone_line(sorted_line, params)
        else
          @array << sorted_line
        end
      end
    end
    
    def set_col_width(hash)
      if hash.class == Hash
        if hash["default"] && hash["default"].respond_to?(:to_f)
          @col_default_width = ((hash["default"].to_f)*10000).to_i.to_f/10000
        end
        hash.keys.each do |key|
          if key.respond_to?(:to_i) && hash[key] && hash[key].respond_to?(:to_f)
            width = ((hash[key].to_f)*10000).to_i.to_f/10000
            pos = key.to_i - 1
            length = col_width.length
            nbr_default = ((pos) - length) < 0 ? 0 : (pos - length)
            nbr_default.times do
              col_width << @col_default_width
            end
            if col_width[pos].nil?
              col_width << width.to_s
            else
              col_width[pos] = width.to_s
            end
          end#if key.respond_to
        end#do key
      end#if hash.class
    end#def
    
    def nbr_col
      nbr = 0
      @array.each do |array|
        nbr = array.length if array.length > nbr
      end
      return nbr
    end
    
    def add_empty(nbr = 1, params = {})
      if params.class == Hash
        @array << [(DefaultStyle.merge(params)).merge({:colspan => nbr})]
      else
        @array << [DefaultStyle.merge({:colspan => nbr})]
      end
    end
    
    def skip(nbr = 1)
      nbr.times do
        @array << [DefaultStyle]
      end
    end
    
    #this method will set the default styles:
    def set_default(hash)
      @default_style = @default_style.merge(hash)
    end
    
    private
    def sort_hash(hash, sort_array)
      new = []
      sort_array.each do |v|
        val = hash[v.to_sym].nil? ? hash[v.to_s] : hash[v.to_sym]
        new << val
      end
      new
    end
    
    def add_clone_line(element, hash)
      res = []
      array = (element.class == Hash) ? element.values : element
      array.each do |val|
        if val.class == Hash
          res << (hash.merge(val))
        elsif val.class == Array
          value = val[0].respond_to?(:to_s) ? val[0].to_s : ""
          params = (val[1].class == Hash) ? val[1] : {}
          res << (hash.merge({:value => value})).merge(params)
        elsif val.respond_to?(:to_s)
          res << (hash.merge({:value => val.to_s}))
        end
      end
      @array<<res
    end
  end#class
  
  class Workbook
    Vertical = ["center", "top", "bottom"]
    Horizontal = ["left", "right", "justified", "center"]
    Underline = ['single','double']
    Font = ['Arial']
    BorderStyle = ['continuous', 'double', 'dashed', 'dotted']
    DefaultStyle = {:colspan => nil, :rowspan => nil, :horizontal => "left", :vertical => "center", :bold => false, :italic => false, :underline => "none", :size => 12, :font => "Arial", :back_color => "none", :color => "#000000", :border => {:left => false, :right => false, :top => false, :bottom => false, :color => "#000000", :style => "continuous", :weight => 1}}
    @@default_style = DefaultStyle.clone
    #the styles are horizontal-vertical-bold-italic-underline-size-font-back_color-color-is_border-border_left-border_right-border_top-border_bottom-border_color-border_style-border_weight"
    #ex: "left-top-true-false-double-12-Arial-#000000-#000000-true-true-true-false-false-#000000-continuous-0
    
    def initialize
      @worksheets = []
      @last_line = []
      @current_line = []
      @nbr_of_worksheet = 0
    end
    
    def add_worksheet(sheetname = "Worksheet")
      @nbr_of_worksheet += 1
      worksheet = Worksheet.new(@nbr_of_worksheet.to_s + "-" + sheetname)
      @worksheets << worksheet
      return worksheet
    end
    
    def build
      buffer = ""
      xml = Builder::XmlMarkup.new(buffer)
      xml.instruct! :xml, :version=>"1.0", :encoding=>"UTF-8" 
      xml.Workbook({
          'xmlns'      => "urn:schemas-microsoft-com:office:spreadsheet", 
          'xmlns:o'    => "urn:schemas-microsoft-com:office:office",
          'xmlns:x'    => "urn:schemas-microsoft-com:office:excel",    
          'xmlns:html' => "http://www.w3.org/TR/REC-html40",
          'xmlns:ss'   => "urn:schemas-microsoft-com:office:spreadsheet" 
        }) do
      
        xml.Styles do
          #We initialize all the styles we will need. And only those that we will need.
          get_styles.each do |style|
            #style is a string. Each string represent the styles separates with "-".
            #the styles are horizontal-vertical-bold-italic-underline-size-font-back_color-color-is_border-border_left-border_right-border_top-border_bottom-border_color-border_style-border_weight"
            #ex: "left-top-true-false-double-12-Arial-#000000-#000000-true-true-true-false-false-#000000-continuous-0
            xml.Style 'ss:ID' => style do
              hash = Hash.new
              array = style.split("-")
              hash[:horizontal] = array[0].capitalize
              hash[:vertical] = array[1].capitalize
              hash[:bold] = (array[2] == "true") ? "1" : "0"
              hash[:italic] = (array[3] == "true") ? "1" : "0"
              hash[:underline] = array[4].capitalize
              hash[:size] = array[5].to_s
              hash[:font] = array[6]
              hash[:back_color] = array[7] != "none" ? array[7] : false
              hash[:color] = array[8]
              hash[:is_border] = (array[9] == "true") ? true : false
              hash[:border_left] = (array[10] == "true") ? true : false
              hash[:border_right] = (array[11] == "true") ? true : false
              hash[:border_top] = (array[12] == "true") ? true : false
              hash[:border_bottom] = (array[13] == "true") ? true : false
              hash[:border_color] = array[14]
              hash[:border_style] = array[15].capitalize
              hash[:border_weight] = array[16].to_s
              
              xml.Alignment 'ss:Horizontal' => hash[:horizontal], 'ss:Vertical' => hash[:vertical]
              xml.Font 'ss:Bold' => hash[:bold], 'ss:Italic' => hash[:italic], 'ss:Underline' => hash[:underline], 'ss:Size' => hash[:size], 'ss:FontName' => hash[:font], 'ss:Color' => hash[:color]
              if hash[:back_color]
                xml.Interior 'ss:Color' => hash[:back_color], 'ss:Pattern' => "Solid"
              end
              if hash[:is_border]
                xml.Borders do
                  if hash[:border_left]
                    xml.Border 'ss:Position' => "Left", 'ss:LineStyle' => hash[:border_style], 'ss:Weight' => hash[:border_weight], 'ss:Color' => hash[:border_color]
                  end
                  if hash[:border_right]
                    xml.Border 'ss:Position' => "Right", 'ss:LineStyle' => hash[:border_style], 'ss:Weight' => hash[:border_weight], 'ss:Color' => hash[:border_color]
                  end
                  if hash[:border_top]
                    xml.Border 'ss:Position' => "Top", 'ss:LineStyle' => hash[:border_style], 'ss:Weight' => hash[:border_weight], 'ss:Color' => hash[:border_color]
                  end
                  if hash[:border_bottom]
                    xml.Border 'ss:Position' => "Bottom", 'ss:LineStyle' => hash[:border_style], 'ss:Weight' => hash[:border_weight], 'ss:Color' => hash[:border_color]
                  end
                end#xml.borders
              end#if hash[:is_border]
            end#xml.style
          end#get_styles.each
        end#xml.styles do
         for object in @worksheets
           xml << worksheetFromArray(object)
         end
      end
      return xml.target! 
    end#build
    
    #this method will set the default styles:
    def set_default(hash)
      @@default_style = @@default_style.merge(hash)
    end
    #######
    private
    #######
    def worksheetFromArray(object)
      buffer =""
      xm = Builder::XmlMarkup.new(buffer) # stream to the text buffer
      name = (object.name.nil? || object.name.empty?) ? "worksheet" : object.name
      xm.Worksheet 'ss:Name' => name do
        xm.Table do
          #Let's add the col_width
          nbr_col = object.nbr_col
          nbr_width = object.col_width.length
          if nbr_width > nbr_col
            object.col_width = object.col_width[0,nbr_col]
          elsif nbr_width < nbr_col
            (nbr_col - nbr_width).times do
              object.col_width << object.col_default_width
            end
          end
          for width in object.col_width
            xm.Column 'ss:Width' => width
          end
          #rows:
          for row in object.array
            xm.Row do
              pos = 0
              for value in (row.class == Hash ? row.values : row)
                can_increment = true
                case value
                  when Hash || Array
                    #let's get the arguments:
                    #get_arguments will also create the instance variable: @value
                    arguments = get_arguments(value, object, pos)
                    xm.Cell arguments do
                      xm.Data @value, 'ss:Type' => 'String'
                    end#xm.Cell do
                  else 
                    if value.respond_to?(:to_s)
                      arguments = get_arguments(@@default_style, object, pos)
                      xm.Cell arguments do
                      xm.Data value.to_s, 'ss:Type' => 'String'
                      end#xm.Cell do
                    else
                      can_increment=false
                    end
                 end#case
                 if can_increment
                  pos+=1
                 end
              end#for value in
              current_become_last
            end#xm.Row do
          end#for row
        end #table
      end #worksheet do
      return xm.target!  # retrieves the buffer
    end#def 
    
    def get_arguments(value, worksheet, pos = nil)
      if value.class == Array
        @value = value[0].respond_to?(:to_s) ? value[0].to_s : ""
        params = (value[1].class == Hash) ? value[1] : {}
      elsif value.class == Hash
        @value = value[:value] ? value[:value].to_s : ""
        params = value
      end
      args = Hash.new
      args['ss:MergeAcross'] = (params[:colspan].to_i)-1 if params[:colspan]
      col = args['ss:MergeAcross'].nil? ? 0 : args['ss:MergeAcross']
      args['ss:MergeDown'] = (params[:rowspan].to_i)-1 if params[:rowspan]
      row = args['ss:MergeDown'].nil? ? 0 : args['ss:MergeDown']
      #let's get the styles:
      args['ss:StyleID'] = get_arguments_name(params, worksheet)
      args['ss:Index'] = get_index(pos)
      @current_line << {:row => row, :col => col, :index => args['ss:Index']}
      args      
    end#def get_arguments
    
    def get_styles
      styles = Array.new
      for object in @worksheets
        tab = object.array
        for row in tab
          for value in (row.class == Hash ? row.values : row)
            case value
              when Hash
                styles << get_arguments_name(value, object)
              when String
              else
            end#case
          end#for value in
        end#for row 
      end#for object
      #on ajoute le style par default:
      styles << get_arguments_name(@@default_style)
      styles.uniq
    end#end get_styles
    
    def is_boolean?(arg)
      res = ([true, false].include? arg) ? true : false
    end
    
    def get_arguments_name(value, worksheet = nil)
      default_style = (worksheet && worksheet.default_style) ? @@default_style.merge(worksheet.default_style) : @@default_style
      horizontal = (Horizontal.include? value[:horizontal]) ? value[:horizontal] : default_style[:horizontal]
      vertical = (Vertical.include? value[:vertical]) ? value[:vertical] : default_style[:vertical]
      bold = (is_boolean?(value[:bold])) ? value[:bold] : default_style[:bold]
      italic = (is_boolean?(value[:italic])) ? value[:italic] : default_style[:italic]
      underline = (Underline.include? value[:underline]) ? value[:underline] : default_style[:underline]
      size = ((2..999).include? value[:size]) ? value[:size].to_i : default_style[:size]
      font = (Font.include? value[:font]) ? value[:font] : default_style[:font]
      back_color = (/#[0-9a-fA-F]{6}/.match(value[:back_color])) ? /#[0-9a-fA-F]{6}/.match(value[:back_color])[0].downcase : default_style[:back_color]
      color = (/#[0-9a-fA-F]{6}/.match(value[:color])) ? /#[0-9a-fA-F]{6}/.match(value[:color])[0].downcase : default_style[:color]
      borders = ((!value[:border].nil?) && (value[:border].class == Hash)) ? DefaultStyle[:border].merge(value[:border]) : DefaultStyle[:border].merge(default_style[:border])
      border_left = (is_boolean?(borders[:left])) ? borders[:left] : false
      border_right = (is_boolean?(borders[:right])) ? borders[:right] : false
      border_top = (is_boolean?(borders[:top])) ? borders[:top] : false
      border_bottom = (is_boolean?(borders[:bottom])) ? borders[:bottom] : false
      border_color = (/#[0-9a-fA-F]{6}/.match(borders[:color])) ? /#[0-9a-fA-F]{6}/.match(borders[:color])[0].downcase : "#000000"
      border_style = (BorderStyle.include? borders[:style]) ? borders[:style] : "continuous"
      border_weight = ((0..10).include? borders[:weight]) ? borders[:weight] : 1
      is_border = border_left || border_right || border_top || border_bottom
      return  "#{horizontal}-#{vertical}-#{bold}-#{italic}-#{underline}-#{size}-#{font}-#{back_color}-#{color}-#{is_border}-#{border_left}-#{border_right}-#{border_top}-#{border_bottom}-#{border_color}-#{border_style}-#{border_weight}"
    end
    
    def get_index(pos)
      default = {:row => 0, :col => 0, :index => 0}
      previous_value = @current_line[positive(pos-1)].nil? ? default : @current_line[positive(pos-1)]
      next_index = previous_value[:index] + previous_value[:col] + 1
      
      #Here is the main calculation of the index of the current cell:
      #index = next_index + (binary(previous_line[:row])*(previous_line[:col]+1))
      previous_line = get_previous_line(next_index)
      while (binary(previous_line[:row])*(previous_line[:col]+1)) > 0 
        next_index += (binary(previous_line[:row])*(previous_line[:col]+1))
        previous_line = get_previous_line(next_index)
      end
      return next_index
    end
    
    def get_previous_line(index)
      #We will search for a hash in last_line that has :index eaqual to index. It may has none.
      default = {:row => 0, :col => 0, :index => 0}
      result = @last_line.select {|val| val[:index] == index}
      return result[0].nil? ? default : result[0]
    end
    
    def current_become_last
      pos = 0
      max = (((@last_line.last)&&(@last_line.last[:index] > @current_line.last[:index]))) ? @last_line.last[:index] : @current_line.last[:index]
      temp_line = []
      max.times do |a|
        if (current = is_in?(@current_line, (a+1)))
          temp_line << current
        elsif (last = is_in?(@last_line, (a+1)))
          if last[:row]>1
            last[:row]-=1
            temp_line << last
          end
        end
      end#max_times do |a|
      @last_line = temp_line
      @current_line = Array.new  
    end
    
    def positive(val)
      return val<0 ? 0 : val
    end
    
    def binary(val)
      return (val == 0) ? 0 : 1
    end
    
    def is_in?(array, index)
      return array.select {|v| v[:index] == index}[0]
    end
  end#class
end

