# frozen_string_literal: true

module Axlsx
  # Table
  # @note Worksheet#add_pivot_table is the recommended way to create tables for your worksheets.
  # @see README for examples
  class PivotTableCacheDefinition
    include Axlsx::OptionsParser

    # Creates a new PivotTable object
    # @param [String] pivot_table The pivot table this cache definition is in
    def initialize(pivot_table)
      @pivot_table = pivot_table
    end

    # # The reference to the pivot table data
    # # @return [PivotTable]
    attr_reader :pivot_table

    # The index of this chart in the workbooks charts collection
    # @return [Integer]
    def index
      pivot_table.sheet.workbook.pivot_tables.index(pivot_table)
    end

    # The part name for this table
    # @return [String]
    def pn
      format(PIVOT_TABLE_CACHE_DEFINITION_PN, index + 1)
    end

    # The identifier for this cache
    # @return [Integer]
    def cache_id
      index + 1
    end

    # The relationship id for this pivot table cache definition.
    # @see Relationship#Id
    # @return [String]
    def rId
      pivot_table.relationships.for(self).Id
    end

    # Serializes the object
    # @param [String] str
    # @return [String]
    def to_xml_string(str = +'')
      str << '<?xml version="1.0" encoding="UTF-8"?>'
      str << '<pivotCacheDefinition xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '" invalid="1" refreshOnLoad="1" recordCount="0">'
      str << '<cacheSource type="worksheet">'
      str << '<worksheetSource ref="' << pivot_table.range << '" sheet="' << pivot_table.data_sheet.name << '"/>'
      str << '</cacheSource>'
      str << '<cacheFields count="' << pivot_table.header_and_filter_cells_count.to_s << '">'
      pivot_table.header_cells.each do |cell|
        str << '<cacheField name="' << cell.clean_value << '" numFmtId="0">'
        if pivot_table.filter_all_values && pivot_table.filter_all_values.has_key?(cell.clean_value)
          values = pivot_table.filter_all_values[cell.clean_value]
          str <<     %{<sharedItems count="#{values.size}">}
           
          values.each do |value|
            str <<      %{<n v="#{value.to_s}" />}
          end
          str <<     '</sharedItems>'
        else
          str <<     '<sharedItems count="0">'
          str <<     '</sharedItems>'
        end
        str << '</cacheField>'
      end
      
      str << '</cacheFields>'
      str << '</pivotCacheDefinition>'
    end
  end
end
