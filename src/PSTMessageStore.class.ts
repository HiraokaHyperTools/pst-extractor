import { PropertyFinder } from './PAUtil'
import { PLNode } from './PLNode'
import { PLSubNode } from './PLSubNode'
import { PSTFile } from './PSTFile.class'
import { PSTObject } from './PSTObject.class'
import { createPUNodeFrom, PUNode } from './PUNode'
import { RootProvider } from './RootProvider'

export class PSTMessageStore extends PSTObject {
  /**
   * Creates an instance of PSTMessageStore.
   * Not much use other than to get the "name" of the PST file.
   * @internal
   * @param {PSTFile} rootProvider
   * @param {DescriptorIndexNode} descriptorIndexNode
   * @memberof PSTMessageStore
   */
  constructor(
    rootProvider: RootProvider,
    node: PLNode,
    subNode: PLSubNode,
    propertyFinder: PropertyFinder
  ) {
    super(rootProvider, node, subNode, propertyFinder)
  }

  /**
   * Requests access to the user node of the internal PST structure.
   */
  public async requestAccessToUserNode(): Promise<PUNode | undefined> {
    return createPUNodeFrom(this._node);
  }
}
