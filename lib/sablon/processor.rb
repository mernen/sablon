# -*- coding: utf-8 -*-
module Sablon
  module Processor
    RELATIONSHIPS_NS_URI = 'http://schemas.openxmlformats.org/package/2006/relationships'
    PICTURE_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    MAIN_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

    def self.process_rels(xml_node, images)
      next_id = next_rel_id(xml_node)
      relationships = xml_node.at_xpath('r:Relationships', 'r' => RELATIONSHIPS_NS_URI)
      images.each do |image|
        relationships.add_child("<Relationship Id='rId#{next_id}' Type='#{IMAGE_TYPE}' Target='media/#{image.name}'/>")
        image.rid = next_id
        next_id += 1
      end
      xml_node
    end

    def self.next_rel_id(xml_node)
      max = 0
      xml_node.xpath('r:Relationships/r:Relationship', 'r' => RELATIONSHIPS_NS_URI).each do |n|
        id = n.attributes['Id'].to_s[3..-1].to_i
        max = id if id > max
      end
      max + 1
    end

    def self.remove_final_blank_page(xml_node)
      children = xml_node.xpath('/w:document/w:body/*')
      found_last = false
      children.reverse.each do |child|
        if found_last
          if child.name == 'p' && child.namespace.prefix == 'w'
            page_break = child.xpath("w:r/w:br[@w:type='page']")
            page_break.remove unless page_break.nil?
            break
          end
        elsif child.name == 'sectPr' && child.namespace.prefix == 'w'
          found_last = true
        end
      end
      xml_node
    end
  end
end
