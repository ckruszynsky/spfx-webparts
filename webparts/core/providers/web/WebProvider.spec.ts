import 'jest';

import { MockWebCommandExecutor } from './executor/MockWebCommandExecutor';
import WebProvider from './WebProvider';

describe("Provider: Web", () => {
  beforeEach(() => {});
  afterEach(() => {});

  it("should get 1 list", async () => {
    //arrange
    const service = new WebProvider(
      new MockWebCommandExecutor({
        lists: [
          {
            Title: "Mocked List",
            DefaultViewUrl: "/mockedLists"
          }
        ]
      })
    );

    //act
    const lists = await service.GetLists();

    //assert
    expect(lists.length).toBe(1);
  });

  it("should get 2 list", async () => {
    //arrange
    const service = new WebProvider(
      new MockWebCommandExecutor({
        lists: [
          {
            Title: "Mocked List",
            DefaultViewUrl: "/mockedLists"
          },
          {
            Title: "Another List",
            DefaultViewUrl: "/mockedLists2"
          }
        ]
      })
    );

    //act
    const lists = await service.GetLists();

    //assert
    expect(lists.length).toBe(2);
  });

  it("should return empty when no lists exists", async () => {
    //arrange
  const service = new WebProvider(
    new MockWebCommandExecutor({lists: []}));
  
    //act  
    const lists = await service.GetLists();

    //assert 
    expect(lists.length).toBe(0);
});
});
